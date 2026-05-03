import * as React from 'react';
import {
  Stack,
  Text,
  TextField,
  PrimaryButton,
  DefaultButton,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  getTheme
} from '@fluentui/react';
import { useMcpConnection } from '@ferrarirosso/mcp-spfx';

import type { IListsChatProps } from './IListsChatProps';
import { runAgent, type IChatMessage, type AgentTraceEntry } from '../agent/runAgent';
import { MessageList } from './MessageList';
import { TracePane } from './TracePane';

const LISTS_SERVER_ID = 'mcp_SharePointRemoteServer';

export const ListsChat: React.FC<IListsChatProps> = (props) => {
  const theme = getTheme();

  // ── MCP connection ──────────────────────────────────────────
  const connection = useMcpConnection({
    environmentId: props.environmentId || undefined,
    serverId: LISTS_SERVER_ID,
    tokenProvider: props.tokenProvider
  });

  // ── Chat state ──────────────────────────────────────────────
  const [history, setHistory] = React.useState<IChatMessage[]>([]);
  const [trace, setTrace] = React.useState<AgentTraceEntry[]>([]);
  const [input, setInput] = React.useState<string>('');
  const [isThinking, setIsThinking] = React.useState<boolean>(false);
  const [statusText, setStatusText] = React.useState<string>('thinking…');
  const [runtimeError, setRuntimeError] = React.useState<string | undefined>(undefined);

  // ── Helpers ─────────────────────────────────────────────────
  const getUserToken = React.useCallback(async (): Promise<string> => {
    if (!props.tokenProvider) {
      throw new Error('AadTokenProvider is not available');
    }
    if (!props.backendApiResource) {
      throw new Error('Backend API resource is not configured in the property pane');
    }
    return props.tokenProvider.getToken(props.backendApiResource);
  }, [props.tokenProvider, props.backendApiResource]);

  const canSend =
    !!props.environmentId &&
    !!props.backendUrl &&
    !!props.backendApiResource &&
    connection.status === 'connected' &&
    !isThinking &&
    input.trim().length > 0;

  const onSubmit = React.useCallback(async () => {
    if (!canSend || !connection.client) return;
    const userMessage = input.trim();
    setInput('');
    setRuntimeError(undefined);
    setIsThinking(true);
    setStatusText('thinking…');

    // Optimistic user-message render so the operator sees their input
    // immediately, not 15+ seconds later when the loop returns. runAgent
    // appends the same {role:'user', content} shape internally, so the
    // returned history matches and there's no flicker.
    const previousHistory = history;
    setHistory((prev) => [...prev, { role: 'user', content: userMessage }]);

    const handleTrace = (entry: AgentTraceEntry): void => {
      setTrace((prev) => [...prev, entry]);
      // Drive the status indicator from trace events. Rate-limit retries
      // get a "throttled, retrying in Ns" message instead of "thinking…"
      // so the user knows what's happening during the backoff.
      if (entry.kind === 'rate_limit_retry') {
        const sourceLabel =
          entry.source === 'azure_openai' ? 'Azure OpenAI' : 'proxy quota';
        setStatusText(
          `throttled by ${sourceLabel} · retrying in ${entry.retryAfterSeconds}s ` +
            `(attempt ${entry.attempt}/${entry.maxAttempts})…`
        );
      } else if (entry.kind === 'tool_call') {
        setStatusText(`calling ${entry.name}…`);
      } else if (entry.kind === 'llm_call') {
        setStatusText('thinking…');
      }
    };

    // Tell the LLM what site it's on. Without this the model defaults to
    // siteId: "root" when the user says "this site" and ends up listing
    // tenant-root system lists instead of the lists in the page's actual
    // SharePoint context.
    const systemPrompt = [
      'You are an assistant embedded in a SharePoint page. You have access to tools',
      'that read and write SharePoint lists for the signed-in user.',
      '',
      'Current SharePoint context (use this when the user refers to "this site",',
      '"the current site", or does not specify a site):',
      `- Site URL: ${props.currentSiteUrl}`,
      `- siteId:   ${props.currentSiteId}`,
      '',
      'When a tool needs a siteId argument and the user has not named a different',
      'site by name or URL, use the siteId above. Pick tools when they help answer',
      'the user\'s question. Always cite the list names or item fields you saw.',
      'When a tool call fails, explain the error plainly.'
    ].join('\n');

    try {
      const updated = await runAgent(
        userMessage,
        previousHistory,
        {
          client: connection.client,
          tools: connection.tools,
          proxyUrl: props.backendUrl,
          getUserToken,
          systemPrompt
        },
        handleTrace
      );
      setHistory(updated);
    } catch (err) {
      const e = err as Error & { code?: string; status?: number };
      const friendly =
        e.code === 'rate_limited'
          ? e.message + ' (The agent retried automatically with backoff but the deployment is still throttling. Bump the model deployment\'s SKU capacity in Azure if this happens often.)'
          : e.message;
      setRuntimeError(friendly);
    } finally {
      setIsThinking(false);
    }
  }, [canSend, connection.client, connection.tools, input, history, props.backendUrl, props.currentSiteId, props.currentSiteUrl, getUserToken]);

  const onReset = React.useCallback(() => {
    setHistory([]);
    setTrace([]);
    setRuntimeError(undefined);
  }, []);

  const onClearTrace = React.useCallback(() => {
    setTrace([]);
  }, []);

  // ── Configuration checks ────────────────────────────────────
  if (!props.environmentId) {
    return (
      <MessageBar messageBarType={MessageBarType.warning}>
        Set the Environment ID in the property pane to connect.
      </MessageBar>
    );
  }

  const missingBackend = !props.backendUrl || !props.backendApiResource;
  if (missingBackend) {
    return (
      <MessageBar messageBarType={MessageBarType.info}>
        Set the LLM backend URL and API resource in the property pane.
        Deploy the backend with <code>cd backend && npm run deploy</code> — the summary
        prints all three values.
      </MessageBar>
    );
  }

  // ── Render ──────────────────────────────────────────────────
  return (
    <Stack tokens={{ childrenGap: 12 }} style={{ padding: 12 }}>
      <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
        <Text variant="xLarge" style={{ fontWeight: 600 }}>
          MCP365 Chat — SharePoint Lists
        </Text>
        <DefaultButton text="Reset" onClick={onReset} disabled={isThinking} />
      </Stack>

      {connection.status === 'connecting' && (
        <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
          <Spinner size={SpinnerSize.small} />
          <Text>Connecting to {LISTS_SERVER_ID}…</Text>
        </Stack>
      )}

      {connection.status === 'error' && connection.error && (
        <MessageBar messageBarType={MessageBarType.error}>
          MCP connection failed: {connection.error.message}
        </MessageBar>
      )}

      {runtimeError && (
        <MessageBar
          messageBarType={MessageBarType.error}
          onDismiss={() => setRuntimeError(undefined)}
        >
          {runtimeError}
        </MessageBar>
      )}

      {connection.status === 'connected' && (
        <>
          <Text
            variant="small"
            style={{ color: theme.palette.neutralSecondary }}
          >
            Connected · {connection.serverInfo?.name} v{connection.serverInfo?.version} ·{' '}
            {connection.tools.length} tools available
          </Text>

          <Stack tokens={{ childrenGap: 12 }} style={{ minWidth: 0 }}>
            <div
              style={{
                minHeight: 240,
                maxHeight: 460,
                overflowY: 'auto',
                padding: 10,
                border: `1px solid ${theme.palette.neutralLight}`,
                borderRadius: 4,
                background: theme.palette.white
              }}
            >
              <MessageList messages={history} isThinking={isThinking} statusText={statusText} />
            </div>

            <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="end">
              <Stack.Item grow>
                <TextField
                  multiline
                  autoAdjustHeight
                  rows={2}
                  value={input}
                  placeholder="Ask about your SharePoint lists… e.g. 'what lists are on this site?'"
                  onChange={(_, v) => setInput(v || '')}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter' && !e.shiftKey) {
                      e.preventDefault();
                      if (canSend) void onSubmit();
                    }
                  }}
                  disabled={isThinking}
                />
              </Stack.Item>
              <PrimaryButton
                text="Send"
                onClick={() => void onSubmit()}
                disabled={!canSend}
              />
            </Stack>

            {props.showTracePane && (
              <TracePane trace={trace} onClear={onClearTrace} />
            )}
          </Stack>
        </>
      )}
    </Stack>
  );
};
