import * as React from 'react';
import {
  Stack,
  Text,
  TextField,
  PrimaryButton,
  DefaultButton,
  MessageBar,
  MessageBarType,
  getTheme
} from '@fluentui/react';

import type { IFoundryChatProps } from './IFoundryChatProps';
import { MessageList } from './MessageList';

export interface IChatMessage {
  role: 'system' | 'user' | 'assistant';
  content: string;
}

const DEFAULT_SYSTEM_PROMPT =
  'You are a brief, friendly assistant embedded in a SharePoint page. Keep replies short.';

async function callProxy(
  proxyUrl: string,
  bearerToken: string,
  messages: IChatMessage[]
): Promise<string> {
  const url = `${proxyUrl.replace(/\/$/, '')}/chat/completions`;
  const res = await fetch(url, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${bearerToken}`
    },
    body: JSON.stringify({ messages })
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Proxy ${res.status}: ${text.slice(0, 500)}`);
  }
  const data = await res.json() as {
    choices?: Array<{ message?: { content?: string } }>;
  };
  const content = data.choices?.[0]?.message?.content;
  if (!content) throw new Error('Proxy returned no content');
  return content;
}

export const FoundryChat: React.FC<IFoundryChatProps> = (props) => {
  const theme = getTheme();

  const [history, setHistory] = React.useState<IChatMessage[]>([]);
  const [input, setInput] = React.useState<string>('');
  const [isThinking, setIsThinking] = React.useState<boolean>(false);
  const [runtimeError, setRuntimeError] = React.useState<string | undefined>(undefined);

  const missingBackend = !props.backendUrl || !props.backendApiResource;

  const canSend = !missingBackend && !isThinking && input.trim().length > 0;

  const onSubmit = React.useCallback(async (): Promise<void> => {
    if (!canSend) return;
    if (!props.tokenProvider) {
      setRuntimeError('AadTokenProvider is not available.');
      return;
    }
    const userText = input.trim();
    setInput('');
    setRuntimeError(undefined);
    setIsThinking(true);

    const baseHistory = history.length
      ? history
      : [{ role: 'system', content: DEFAULT_SYSTEM_PROMPT } as IChatMessage];
    const next: IChatMessage[] = [...baseHistory, { role: 'user', content: userText }];
    setHistory(next);

    try {
      const token = await props.tokenProvider.getToken(props.backendApiResource);
      const reply = await callProxy(props.backendUrl, token, next);
      setHistory([...next, { role: 'assistant', content: reply }]);
    } catch (err) {
      setRuntimeError((err as Error).message);
      setHistory(history);
    } finally {
      setIsThinking(false);
    }
  }, [canSend, input, history, props.backendUrl, props.backendApiResource, props.tokenProvider]);

  const onReset = React.useCallback((): void => {
    setHistory([]);
    setRuntimeError(undefined);
  }, []);

  if (missingBackend) {
    return (
      <MessageBar messageBarType={MessageBarType.info}>
        Set the backend URL and API resource in the property pane.
        Deploy the backend with <code>npm run deploy</code> from this webpart — the
        property pane is then auto-wired by <code>npm run setup</code>.
      </MessageBar>
    );
  }

  return (
    <Stack tokens={{ childrenGap: 12 }} style={{ padding: 12 }}>
      <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
        <Stack>
          <Text variant="xLarge" style={{ fontWeight: 600 }}>
            MCP365 Chat — Foundry Showcase
          </Text>
          <Text variant="small" style={{ color: theme.palette.neutralSecondary }}>
            Plain chat against the protected proxy. If this answers, your auth chain works.
          </Text>
        </Stack>
        <DefaultButton text="Reset" onClick={onReset} disabled={isThinking} />
      </Stack>

      {runtimeError && (
        <MessageBar
          messageBarType={MessageBarType.error}
          onDismiss={() => setRuntimeError(undefined)}
        >
          {runtimeError}
        </MessageBar>
      )}

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
        <MessageList messages={history} isThinking={isThinking} />
      </div>

      <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="end">
        <Stack.Item grow>
          <TextField
            multiline
            autoAdjustHeight
            rows={2}
            value={input}
            placeholder="Ask anything… e.g. 'summarise the role of SharePoint workspaces in 3 bullets'"
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
        <PrimaryButton text="Send" onClick={() => void onSubmit()} disabled={!canSend} />
      </Stack>
    </Stack>
  );
};
