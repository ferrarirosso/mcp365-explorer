import * as React from 'react';
import {
  PrimaryButton,
  DefaultButton,
  TextField,
  Stack,
  Text,
  Spinner,
  SpinnerSize,
  getTheme
} from '@fluentui/react';
import type { McpBrowserClient, IMcpCallResult, McpLogHandler } from '../services/McpBrowserClient';

interface IShowcaseProps {
  client: McpBrowserClient;
  theme: ReturnType<typeof getTheme>;
  onLog: McpLogHandler;
}

interface ISearchResult {
  reply: string;
  messageCount: number;
}

interface IShowcaseState {
  searchQuery: string;
  searchResult: ISearchResult | undefined;
  // Send email
  sendTo: string;
  sendSubject: string;
  sendBody: string;
  sendResult: string | undefined;
  loading: string | undefined;
  error: string | undefined;
}

function extractJsonFromContent(result: IMcpCallResult): unknown {
  if (!result.content || result.content.length === 0) return undefined;
  const text = result.content[0].text;
  if (!text) return undefined;
  const jsonStart = text.indexOf('{');
  const arrStart = text.indexOf('[');
  const start = jsonStart === -1 ? arrStart : (arrStart === -1 ? jsonStart : Math.min(jsonStart, arrStart));
  if (start === -1) return undefined;
  try {
    const outer = JSON.parse(text.substring(start));
    if (outer && typeof outer === 'object' && typeof outer.response === 'string') {
      try { return JSON.parse(outer.response); } catch { return outer; }
    }
    return outer;
  } catch { return undefined; }
}

function extractSearchResult(result: IMcpCallResult): ISearchResult | undefined {
  const parsed = extractJsonFromContent(result);
  if (!parsed || typeof parsed !== 'object') return undefined;
  const obj = parsed as Record<string, unknown>;

  // The response has a 'reply' field with markdown summary and 'messageIds' array
  const reply = obj.reply as string | undefined;
  const messageIds = obj.messageIds as string[] | undefined;

  if (reply) {
    // Clean up escaped newlines
    const cleanReply = reply.replace(/\\n/g, '\n').replace(/\\"/g, '"');
    return {
      reply: cleanReply,
      messageCount: messageIds ? messageIds.length : 0
    };
  }

  // Fallback: check if there's a message field
  const message = obj.message as string | undefined;
  if (message) {
    return { reply: message, messageCount: 0 };
  }

  return undefined;
}

export const MailShowcase: React.FC<IShowcaseProps> = ({ client, theme }) => {
  const [state, setState] = React.useState<IShowcaseState>({
    searchQuery: '',
    searchResult: undefined,
    sendTo: '',
    sendSubject: '',
    sendBody: '',
    sendResult: undefined,
    loading: undefined,
    error: undefined
  });

  const isLoading = !!state.loading;

  const handleSearch = React.useCallback(async (): Promise<void> => {
    if (!state.searchQuery.trim()) return;
    setState(prev => ({ ...prev, loading: 'SearchMessages', error: undefined, searchResult: undefined }));
    try {
      const result = await client.callTool('SearchMessages', { message: state.searchQuery.trim() });
      setState(prev => ({ ...prev, loading: undefined, searchResult: extractSearchResult(result) }));
    } catch (err) { setState(prev => ({ ...prev, loading: undefined, error: (err as Error).message })); }
  }, [client, state.searchQuery]);

  const handleRecent = React.useCallback(async (): Promise<void> => {
    setState(prev => ({ ...prev, loading: 'SearchMessages-recent', error: undefined, searchResult: undefined }));
    try {
      const result = await client.callTool('SearchMessages', { message: 'show me my recent emails' });
      setState(prev => ({ ...prev, loading: undefined, searchResult: extractSearchResult(result) }));
    } catch (err) { setState(prev => ({ ...prev, loading: undefined, error: (err as Error).message })); }
  }, [client]);

  const handleSendEmail = React.useCallback(async (): Promise<void> => {
    if (!state.sendTo.trim() || !state.sendSubject.trim()) return;
    setState(prev => ({ ...prev, loading: 'SendEmailWithAttachments', error: undefined, sendResult: undefined }));
    try {
      const result = await client.callTool('SendEmailWithAttachments', {
        to: state.sendTo.split(',').map(e => e.trim()).filter(e => e !== ''),
        subject: state.sendSubject.trim(),
        body: state.sendBody.trim() || ' '
      });
      const parsed = extractJsonFromContent(result);
      const isError = result.isError || (parsed && typeof parsed === 'object' && (parsed as Record<string, unknown>).Error);
      if (isError) {
        setState(prev => ({ ...prev, loading: undefined, sendResult: undefined, error: result.content?.[0]?.text || 'Send failed' }));
      } else {
        setState(prev => ({ ...prev, loading: undefined, sendResult: 'Email sent successfully' }));
      }
    } catch (err) { setState(prev => ({ ...prev, loading: undefined, error: (err as Error).message })); }
  }, [client, state.sendTo, state.sendSubject, state.sendBody]);

  const cardStyle: React.CSSProperties = {
    border: `1px solid ${theme.palette.neutralLight}`, borderRadius: 8, padding: 12, backgroundColor: theme.palette.white, marginTop: 8
  };

  return (
    <div style={{ marginTop: 12 }}>
      <Text variant="small" styles={{ root: { color: theme.palette.neutralTertiary, marginBottom: 12, display: 'block' } }}>
        Search your mailbox and send emails using MCP tools.
      </Text>

      {state.error && (
        <Text variant="small" styles={{ root: { color: theme.palette.red, marginBottom: 8, display: 'block' } }}>{state.error}</Text>
      )}

      {/* ── Search ───────────────────────────────────────────── */}
      <div style={cardStyle}>
        <Text styles={{ root: { fontWeight: 600, marginBottom: 8, display: 'block' } }}>Search Mailbox (SearchMessages)</Text>
        <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="end">
          <TextField
            placeholder="Natural language search (e.g., emails about the project)"
            value={state.searchQuery}
            onChange={(_, val) => setState(prev => ({ ...prev, searchQuery: val || '' }))}
            styles={{ root: { width: 350 } }}
            onKeyDown={(e) => { if (e.key === 'Enter') { void handleSearch(); } }}
          />
          <PrimaryButton text={state.loading === 'SearchMessages' ? 'Searching...' : 'Search'} iconProps={{ iconName: 'Search' }} onClick={handleSearch} disabled={isLoading || !state.searchQuery.trim()} />
          <DefaultButton text={isLoading ? 'Loading...' : 'Recent'} iconProps={{ iconName: 'Inbox' }} onClick={handleRecent} disabled={isLoading} />
        </Stack>

        {state.searchResult && (
          <div style={{ marginTop: 8 }}>
            {state.searchResult.messageCount > 0 && (
              <Text variant="small" styles={{ root: { fontWeight: 600, color: theme.palette.neutralSecondary, marginBottom: 8, display: 'block' } }}>
                {state.searchResult.messageCount} message{state.searchResult.messageCount !== 1 ? 's' : ''} referenced
              </Text>
            )}
            <div style={{
              padding: 12,
              backgroundColor: theme.palette.neutralLighterAlt,
              borderRadius: 4,
              fontSize: 13,
              lineHeight: '1.6',
              whiteSpace: 'pre-wrap',
              wordBreak: 'break-word',
              maxHeight: 400,
              overflow: 'auto'
            }}
            dangerouslySetInnerHTML={{
              __html: state.searchResult.reply
                .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>')
                .replace(/^## (.*?)$/gm, '<h3 style="margin:12px 0 4px 0;font-size:14px">$1</h3>')
                .replace(/^### (.*?)$/gm, '<h4 style="margin:8px 0 4px 0;font-size:13px">$1</h4>')
                .replace(/^- (.*?)$/gm, '<div style="padding-left:12px">&bull; $1</div>')
                .replace(/\[(\d+)\]\([^)]+\)/g, '<sup style="color:' + theme.palette.themePrimary + '">[$1]</sup>')
                .replace(/---/g, '<hr style="border:none;border-top:1px solid ' + theme.palette.neutralLight + ';margin:8px 0">')
                .replace(/\n\n/g, '<br/>')
                .replace(/\n/g, '<br/>')
            }}
            />
          </div>
        )}
      </div>

      {/* ── Send Email ───────────────────────────────────────── */}
      <div style={cardStyle}>
        <Text styles={{ root: { fontWeight: 600, marginBottom: 8, display: 'block' } }}>Send Email (SendEmailWithAttachments)</Text>
        <Stack tokens={{ childrenGap: 8 }}>
          <TextField
            label="To"
            placeholder="recipient@contoso.com (comma-separated for multiple)"
            value={state.sendTo}
            onChange={(_, val) => setState(prev => ({ ...prev, sendTo: val || '' }))}
          />
          <TextField
            label="Subject"
            placeholder="Email subject"
            value={state.sendSubject}
            onChange={(_, val) => setState(prev => ({ ...prev, sendSubject: val || '' }))}
          />
          <TextField
            label="Body"
            placeholder="Email body"
            value={state.sendBody}
            onChange={(_, val) => setState(prev => ({ ...prev, sendBody: val || '' }))}
            multiline
            rows={3}
          />
          <PrimaryButton
            text={state.loading === 'SendEmailWithAttachments' ? 'Sending...' : 'Send'}
            iconProps={{ iconName: 'Send' }}
            onClick={handleSendEmail}
            disabled={isLoading || !state.sendTo.trim() || !state.sendSubject.trim()}
            styles={{ root: { maxWidth: 120 } }}
          />
          {state.sendResult && (
            <Text variant="small" styles={{ root: { color: '#107c10', fontWeight: 600 } }}>{state.sendResult}</Text>
          )}
        </Stack>
      </div>

      {isLoading && <Spinner size={SpinnerSize.small} label={`Calling ${state.loading}...`} styles={{ root: { marginTop: 8 } }} />}
    </div>
  );
};
