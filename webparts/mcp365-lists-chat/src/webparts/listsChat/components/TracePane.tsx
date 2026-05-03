import * as React from 'react';
import { Stack, Text, IconButton, DefaultButton, getTheme } from '@fluentui/react';
import type { AgentTraceEntry } from '../agent/runAgent';

export interface ITracePaneProps {
  trace: AgentTraceEntry[];
  onClear?: () => void;
}

const STATUS_COLORS: Record<string, string> = {
  llm_call: '#0078d4',
  tool_call: '#107c10',
  rate_limit_retry: '#ca5010',
  error: '#d13438'
};

/**
 * Waterfall-style trace of everything the agent loop did: each LLM call,
 * each MCP tool invocation, each error. Renders the plumbing that the
 * chat bubbles hide.
 *
 * Header includes Copy log + Clear log actions; both are no-ops when there
 * are no entries yet.
 */
export const TracePane: React.FC<ITracePaneProps> = ({ trace, onClear }) => {
  const theme = getTheme();
  const [expanded, setExpanded] = React.useState<number | undefined>(undefined);
  const [copyState, setCopyState] = React.useState<'idle' | 'copied' | 'failed'>('idle');

  const onCopy = React.useCallback((): void => {
    if (!trace.length) return;
    const text = JSON.stringify(trace, null, 2);
    const finish = (state: 'copied' | 'failed'): void => {
      setCopyState(state);
      window.setTimeout(() => setCopyState('idle'), 1500);
    };
    if (typeof navigator !== 'undefined' && navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).then(() => finish('copied'), () => finish('failed'));
      return;
    }
    // Fallback for older Edge / restrictive contexts.
    try {
      const textarea = document.createElement('textarea');
      textarea.value = text;
      textarea.style.position = 'fixed';
      textarea.style.left = '-9999px';
      document.body.appendChild(textarea);
      textarea.select();
      document.execCommand('copy');
      document.body.removeChild(textarea);
      finish('copied');
    } catch {
      finish('failed');
    }
  }, [trace]);

  const copyLabel =
    copyState === 'copied' ? 'Copied' : copyState === 'failed' ? 'Copy failed' : 'Copy log';

  return (
    <Stack
      tokens={{ childrenGap: 6 }}
      style={{
        padding: 12,
        border: `1px solid ${theme.palette.neutralLight}`,
        borderRadius: 4,
        background: theme.palette.neutralLighterAlt
      }}
    >
      <Stack horizontal verticalAlign="center" horizontalAlign="space-between" tokens={{ childrenGap: 8 }}>
        <Text
          variant="small"
          style={{
            textTransform: 'uppercase',
            letterSpacing: 0.6,
            color: theme.palette.neutralSecondary,
            fontWeight: 600
          }}
        >
          Trace ({trace.length})
        </Text>
        <Stack horizontal tokens={{ childrenGap: 6 }}>
          <DefaultButton
            text={copyLabel}
            iconProps={{ iconName: 'Copy' }}
            onClick={onCopy}
            disabled={trace.length === 0}
            styles={{ root: { height: 26, minWidth: 0, padding: '0 10px' }, label: { fontSize: 12 } }}
          />
          <DefaultButton
            text="Clear log"
            iconProps={{ iconName: 'Delete' }}
            onClick={onClear}
            disabled={!onClear || trace.length === 0}
            styles={{ root: { height: 26, minWidth: 0, padding: '0 10px' }, label: { fontSize: 12 } }}
          />
        </Stack>
      </Stack>

      {trace.length === 0 && (
        <Text variant="small" style={{ color: theme.palette.neutralTertiary }}>
          No activity yet. Send a message to see the agent loop.
        </Text>
      )}

      <Stack tokens={{ childrenGap: 6 }} style={{ maxHeight: 320, overflowY: 'auto' }}>
        {trace.map((entry, i) => {
          const isOpen = expanded === i;
          const color = STATUS_COLORS[entry.kind] || theme.palette.neutralSecondary;
          const label = labelFor(entry);
          const duration =
            'durationMs' in entry && typeof entry.durationMs === 'number'
              ? `${entry.durationMs} ms`
              : '';

          return (
            <div
              key={i}
              style={{
                border: `1px solid ${theme.palette.neutralLight}`,
                borderLeft: `3px solid ${color}`,
                borderRadius: 3,
                background: theme.palette.white,
                padding: '6px 8px',
                fontSize: 12
              }}
            >
              <Stack
                horizontal
                verticalAlign="center"
                tokens={{ childrenGap: 6 }}
                style={{ cursor: 'pointer' }}
                onClick={() => setExpanded(isOpen ? undefined : i)}
              >
                <Text
                  variant="small"
                  style={{
                    color,
                    fontFamily: 'Consolas, Menlo, monospace',
                    fontWeight: 600,
                    minWidth: 80
                  }}
                >
                  #{padNum(i + 1, 3)}
                </Text>
                <Text variant="small" style={{ flexGrow: 1 }}>
                  {label}
                </Text>
                <Text variant="small" style={{ color: theme.palette.neutralSecondary }}>
                  {duration}
                </Text>
                <IconButton
                  iconProps={{ iconName: isOpen ? 'ChevronDown' : 'ChevronRight' }}
                  styles={{ root: { height: 20, width: 20 } }}
                />
              </Stack>
              {isOpen && (
                <pre
                  style={{
                    margin: '6px 0 0',
                    padding: 6,
                    background: theme.palette.neutralLighterAlt,
                    borderRadius: 3,
                    fontSize: 11,
                    fontFamily: 'Consolas, Menlo, monospace',
                    whiteSpace: 'pre-wrap',
                    wordBreak: 'break-word',
                    maxHeight: 240,
                    overflow: 'auto'
                  }}
                >
                  {JSON.stringify(entry, null, 2)}
                </pre>
              )}
            </div>
          );
        })}
      </Stack>
    </Stack>
  );
};

function padNum(n: number, width: number): string {
  let s = String(n);
  while (s.length < width) s = '0' + s;
  return s;
}

function labelFor(entry: AgentTraceEntry): string {
  if (entry.kind === 'llm_call') {
    const tokens = entry.tokenUsage?.total_tokens;
    return tokens ? `llm_call (${tokens} tokens)` : 'llm_call';
  }
  if (entry.kind === 'tool_call') {
    return entry.error ? `tool_call ${entry.name} — ERROR` : `tool_call ${entry.name}`;
  }
  if (entry.kind === 'rate_limit_retry') {
    const src = entry.source === 'azure_openai' ? 'AOAI 429' : 'proxy 429';
    return `${src} — retry ${entry.attempt}/${entry.maxAttempts} after ${entry.retryAfterSeconds}s`;
  }
  return `error: ${entry.message}`;
}
