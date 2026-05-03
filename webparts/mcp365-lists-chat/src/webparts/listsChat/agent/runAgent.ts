import type { McpBrowserClient, IMcpTool, IMcpCallResult } from '@ferrarirosso/mcp-browser-client';
import { mcpToolsToOpenAI, mcpResultToToolContent } from './toolTranslation';

// =============================================================================
// Types
// =============================================================================

/** OpenAI-compatible message — the shape we send to the proxy. */
export interface IChatMessage {
  role: 'system' | 'user' | 'assistant' | 'tool';
  content: string | null;
  tool_calls?: IToolCall[];
  tool_call_id?: string;
  name?: string;
}

export interface IToolCall {
  id: string;
  type: 'function';
  function: { name: string; arguments: string };
}

export interface ITokenUsage {
  prompt_tokens?: number;
  completion_tokens?: number;
  total_tokens?: number;
}

export interface IAgentContext {
  client: McpBrowserClient;
  tools: IMcpTool[];
  proxyUrl: string;
  /**
   * Returns a fresh AAD bearer token for the backend API scope on demand.
   * Typically wraps `AadTokenProvider.getToken(backendApiResource)`.
   */
  getUserToken: () => Promise<string>;
  /** Hard cap on turns to prevent runaway loops. Default 10. */
  maxTurns?: number;
  /** Optional system prompt; default added if omitted. */
  systemPrompt?: string;
}

export type AgentTraceEntry =
  | { kind: 'llm_call'; startedAt: number; durationMs: number; tokenUsage?: ITokenUsage }
  | {
      kind: 'tool_call';
      startedAt: number;
      durationMs: number;
      name: string;
      args: unknown;
      result?: IMcpCallResult;
      error?: string;
    }
  | {
      kind: 'rate_limit_retry';
      startedAt: number;
      retryAfterSeconds: number;
      attempt: number;
      maxAttempts: number;
      source: 'azure_openai' | 'proxy_caller_quota';
    }
  | { kind: 'error'; startedAt: number; message: string };

export type AgentTraceHandler = (entry: AgentTraceEntry) => void;

// =============================================================================
// Constants
// =============================================================================

const DEFAULT_MAX_TURNS = 10;
/** How many times we'll wait + retry on a 429 before giving up. */
const MAX_RATE_LIMIT_RETRIES = 2;
/** Cap on backoff so a misbehaving server can't park us forever. */
const MAX_RETRY_AFTER_SECONDS = 60;
/** Used when 429 arrives without a Retry-After header. */
const FALLBACK_RETRY_AFTER_SECONDS = 30;

/**
 * Parse a Retry-After header value. RFC 7231 allows either a delta-seconds
 * integer or an HTTP date; in practice Azure OpenAI sends seconds. Returns
 * undefined if the value is missing or unparseable.
 */
function parseRetryAfterSeconds(headerValue: string | null): number | undefined {
  if (!headerValue) return undefined;
  const trimmed = headerValue.trim();
  const asInt = parseInt(trimmed, 10);
  if (!Number.isNaN(asInt) && asInt >= 0) {
    return Math.min(asInt, MAX_RETRY_AFTER_SECONDS);
  }
  const asDate = Date.parse(trimmed);
  if (!Number.isNaN(asDate)) {
    const delta = Math.ceil((asDate - Date.now()) / 1000);
    return delta > 0 ? Math.min(delta, MAX_RETRY_AFTER_SECONDS) : 0;
  }
  return undefined;
}

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

const DEFAULT_SYSTEM_PROMPT =
  'You are an assistant embedded in a SharePoint page. You have access to tools that read and write SharePoint lists for the signed-in user. ' +
  'Pick tools when they help answer the user\'s question. Always cite the list names or item fields you saw. ' +
  'When a tool call fails, explain the error plainly.';

// =============================================================================
// The loop
// =============================================================================

/**
 * Bare-metal LLM tool-calling loop.
 *
 *   while not done:
 *     1. POST the running conversation + tool definitions to the proxy
 *     2. append the assistant's reply to the conversation
 *     3. if the reply has tool_calls → execute each via MCP, append results
 *        → continue the loop
 *     4. if the reply has plain content → return
 *
 * Returns the full updated message history. Emits a trace entry for every
 * LLM call and every tool call so the caller can render a live waterfall.
 */
export async function runAgent(
  userMessage: string,
  history: IChatMessage[],
  ctx: IAgentContext,
  onTrace: AgentTraceHandler
): Promise<IChatMessage[]> {
  const maxTurns = ctx.maxTurns ?? DEFAULT_MAX_TURNS;
  const openAiTools = mcpToolsToOpenAI(ctx.tools);

  const messages: IChatMessage[] = history.length
    ? history.slice()
    : [{ role: 'system', content: ctx.systemPrompt || DEFAULT_SYSTEM_PROMPT }];

  messages.push({ role: 'user', content: userMessage });

  for (let turn = 0; turn < maxTurns; turn++) {
    // ── 1. Call the LLM via our proxy (with 429-aware retry) ───
    const llmStart = Date.now();
    let choice: { message: IChatMessage; finish_reason?: string };
    let usage: ITokenUsage | undefined;
    let rateLimitRetries = 0;

    try {
      while (true) {
        const userToken = await ctx.getUserToken();
        const res = await fetch(`${ctx.proxyUrl.replace(/\/$/, '')}/chat/completions`, {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            Authorization: `Bearer ${userToken}`
          },
          body: JSON.stringify({ messages, tools: openAiTools, tool_choice: 'auto' })
        });

        // Happy path
        if (res.ok) {
          const data = await res.json() as {
            choices?: Array<{ message: IChatMessage; finish_reason?: string }>;
            usage?: ITokenUsage;
          };
          if (!data.choices || !data.choices.length) {
            throw new Error('Proxy returned no choices');
          }
          choice = data.choices[0];
          usage = data.usage;
          break;
        }

        // 429 — back off and retry up to MAX_RATE_LIMIT_RETRIES times.
        // Distinguishes Azure OpenAI throttling (model deployment RPM/TPM
        // cap) from our own per-caller quota — the proxy is the only one
        // that emits the X-RateLimit-Remaining-Minute header, so its
        // presence tells us "this 429 is from the proxy, not upstream".
        if (res.status === 429 && rateLimitRetries < MAX_RATE_LIMIT_RETRIES) {
          const retryAfter =
            parseRetryAfterSeconds(res.headers.get('retry-after')) ??
            FALLBACK_RETRY_AFTER_SECONDS;
          const isProxyQuota = res.headers.get('x-ratelimit-remaining-minute') === '0';
          rateLimitRetries++;
          onTrace({
            kind: 'rate_limit_retry',
            startedAt: Date.now(),
            retryAfterSeconds: retryAfter,
            attempt: rateLimitRetries,
            maxAttempts: MAX_RATE_LIMIT_RETRIES,
            source: isProxyQuota ? 'proxy_caller_quota' : 'azure_openai'
          });
          // Drain the body to free the connection.
          await res.text().catch(() => undefined);
          await sleep(retryAfter * 1000);
          continue;
        }

        // Non-retryable failure (or 429 retries exhausted).
        const text = await res.text();
        if (res.status === 429) {
          const tried: number = MAX_RATE_LIMIT_RETRIES;
          const err = new Error(
            `Rate limited. Tried ${tried} times with backoff but the deployment is still throttling. Wait a minute and send again.`
          ) as Error & { code?: string; status?: number };
          err.code = 'rate_limited';
          err.status = 429;
          throw err;
        }
        throw new Error(`Proxy ${res.status}: ${text.slice(0, 500)}`);
      }
    } catch (err) {
      const message = (err as Error).message;
      onTrace({ kind: 'error', startedAt: llmStart, message });
      throw err;
    }
    onTrace({
      kind: 'llm_call',
      startedAt: llmStart,
      durationMs: Date.now() - llmStart,
      tokenUsage: usage
    });

    // ── 2. Append assistant reply ──────────────────────────────
    messages.push(choice.message);

    // ── 3. If no tool_calls, we're done ────────────────────────
    const toolCalls = choice.message.tool_calls;
    if (!toolCalls || !toolCalls.length) {
      return messages;
    }

    // ── 4. Execute each tool call via MCP ──────────────────────
    for (const call of toolCalls) {
      const toolStart = Date.now();
      let parsedArgs: Record<string, unknown> = {};
      try {
        parsedArgs = call.function.arguments ? JSON.parse(call.function.arguments) : {};
      } catch {
        // Model produced invalid JSON — surface to the loop so it can self-correct
        const errMsg = `Invalid JSON arguments from model: ${call.function.arguments}`;
        onTrace({
          kind: 'tool_call',
          startedAt: toolStart,
          durationMs: Date.now() - toolStart,
          name: call.function.name,
          args: call.function.arguments,
          error: errMsg
        });
        messages.push({
          role: 'tool',
          tool_call_id: call.id,
          name: call.function.name,
          content: `Error: ${errMsg}`
        });
        continue;
      }

      try {
        const result = await ctx.client.callTool(call.function.name, parsedArgs);
        onTrace({
          kind: 'tool_call',
          startedAt: toolStart,
          durationMs: Date.now() - toolStart,
          name: call.function.name,
          args: parsedArgs,
          result
        });
        messages.push({
          role: 'tool',
          tool_call_id: call.id,
          name: call.function.name,
          content: mcpResultToToolContent(result)
        });
      } catch (err) {
        const message = (err as Error).message;
        onTrace({
          kind: 'tool_call',
          startedAt: toolStart,
          durationMs: Date.now() - toolStart,
          name: call.function.name,
          args: parsedArgs,
          error: message
        });
        messages.push({
          role: 'tool',
          tool_call_id: call.id,
          name: call.function.name,
          content: `Tool call failed: ${message}`
        });
      }
    }
    // Loop back — the LLM now sees the tool results and decides what's next.
  }

  // Hit the turn cap without a plain-content reply. Return what we have;
  // the caller can show a "stopped after N turns" note.
  messages.push({
    role: 'assistant',
    content: `(Stopped after ${maxTurns} turns without a final answer. The model kept requesting tools.)`
  });
  return messages;
}
