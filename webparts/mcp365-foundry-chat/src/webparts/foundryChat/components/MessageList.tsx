import * as React from 'react';
import { Stack, Text, getTheme, Spinner, SpinnerSize } from '@fluentui/react';
import type { IChatMessage } from './FoundryChat';

export interface IMessageListProps {
  messages: IChatMessage[];
  isThinking: boolean;
}

export const MessageList: React.FC<IMessageListProps> = ({ messages, isThinking }) => {
  const theme = getTheme();

  const visible = messages.filter(
    (m) => (m.role === 'user' || m.role === 'assistant') && m.content
  );

  return (
    <Stack tokens={{ childrenGap: 10 }}>
      {visible.map((m, i) => (
        <Stack
          key={i}
          horizontal
          horizontalAlign={m.role === 'user' ? 'end' : 'start'}
        >
          <div
            style={{
              maxWidth: '80%',
              padding: '8px 12px',
              borderRadius: 8,
              background:
                m.role === 'user'
                  ? theme.palette.themeLighter
                  : theme.palette.neutralLighterAlt,
              border: `1px solid ${theme.palette.neutralLight}`,
              whiteSpace: 'pre-wrap',
              wordBreak: 'break-word'
            }}
          >
            <Text
              block
              variant="small"
              style={{
                color: theme.palette.neutralSecondary,
                textTransform: 'uppercase',
                letterSpacing: 0.5,
                marginBottom: 2,
                fontSize: 10
              }}
            >
              {m.role}
            </Text>
            <Text block>{m.content}</Text>
          </div>
        </Stack>
      ))}
      {isThinking && (
        <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
          <Spinner size={SpinnerSize.xSmall} />
          <Text variant="small" style={{ color: theme.palette.neutralSecondary }}>
            thinking…
          </Text>
        </Stack>
      )}
    </Stack>
  );
};
