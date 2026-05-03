import type { IMcpTool, IMcpCallResult } from '@ferrarirosso/mcp-browser-client';

/** OpenAI-compatible function definition shape. */
export interface IOpenAIFunctionTool {
  type: 'function';
  function: {
    name: string;
    description: string;
    parameters: Record<string, unknown>;
  };
}

const DESCRIPTION_MAX = 500;
const EMPTY_SCHEMA: Record<string, unknown> = { type: 'object', properties: {} };

/**
 * Convert a list of MCP tools (from `tools/list`) into OpenAI function-tool
 * definitions that can be dropped into a chat completions `tools` array.
 *
 * Caps descriptions at 500 chars — MCP servers sometimes ship paragraphs,
 * and description bytes are paid on every completion request.
 */
export function mcpToolsToOpenAI(tools: IMcpTool[]): IOpenAIFunctionTool[] {
  return tools.map((t) => ({
    type: 'function',
    function: {
      name: t.name,
      description: (t.description || '').slice(0, DESCRIPTION_MAX),
      parameters:
        t.inputSchema && typeof t.inputSchema === 'object'
          ? (t.inputSchema as Record<string, unknown>)
          : EMPTY_SCHEMA
    }
  }));
}

/**
 * Flatten an MCP tool call result into a plain string for the `tool` role
 * message. MCP returns `content: [{ type: 'text', text: '...' }, ...]`; we
 * concatenate the text parts and stringify anything exotic.
 */
export function mcpResultToToolContent(result: IMcpCallResult): string {
  if (!result || !Array.isArray(result.content)) {
    return JSON.stringify(result ?? {});
  }
  return result.content
    .map((part) => {
      if (part && part.type === 'text' && typeof part.text === 'string') {
        return part.text;
      }
      return JSON.stringify(part);
    })
    .join('\n');
}
