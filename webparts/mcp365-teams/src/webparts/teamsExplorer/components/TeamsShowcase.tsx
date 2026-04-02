import * as React from 'react';
import {
  PrimaryButton,
  DefaultButton,
  TextField,
  Stack,
  Text,
  Spinner,
  SpinnerSize,
  Icon,
  ChoiceGroup,
  getTheme
} from '@fluentui/react';
import type { IChoiceGroupOption } from '@fluentui/react';
import type { McpBrowserClient, IMcpCallResult, McpLogHandler } from '../services/McpBrowserClient';

interface IShowcaseProps {
  client: McpBrowserClient;
  theme: ReturnType<typeof getTheme>;
  onLog: McpLogHandler;
  userEmail: string;
  userId: string;
}

interface ITeamItem {
  displayName: string;
  description: string;
  id: string;
}

interface IChannelItem {
  displayName: string;
  description: string;
  id: string;
}

interface IShowcaseState {
  teams: ITeamItem[];
  selectedTeamId: string | undefined;
  selectedTeamName: string | undefined;
  channels: IChannelItem[];
  selfMessage: string;
  selfMessageResult: string | undefined;
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

function extractArray(result: IMcpCallResult): Array<Record<string, unknown>> {
  const parsed = extractJsonFromContent(result);
  if (!parsed) return [];
  if (Array.isArray(parsed)) return parsed as Array<Record<string, unknown>>;
  if (parsed && typeof parsed === 'object') {
    const obj = parsed as Record<string, unknown>;
    // Check common wrapper keys
    if (Array.isArray(obj.value)) return obj.value as Array<Record<string, unknown>>;
    if (Array.isArray(obj.teams)) return obj.teams as Array<Record<string, unknown>>;
    if (Array.isArray(obj.channels)) return obj.channels as Array<Record<string, unknown>>;
    if (Array.isArray(obj.chats)) return obj.chats as Array<Record<string, unknown>>;
    return [obj];
  }
  return [];
}

export const TeamsShowcase: React.FC<IShowcaseProps> = ({ client, theme, userEmail, userId }) => {
  const [state, setState] = React.useState<IShowcaseState>({
    teams: [],
    selectedTeamId: undefined,
    selectedTeamName: undefined,
    channels: [],
    selfMessage: 'Hello from MCP365 Explorer!',
    selfMessageResult: undefined,
    loading: undefined,
    error: undefined
  });

  const isLoading = !!state.loading;

  const cardStyle: React.CSSProperties = {
    border: `1px solid ${theme.palette.neutralLight}`, borderRadius: 8, padding: 12, backgroundColor: theme.palette.white, marginTop: 8
  };

  // ── List Teams ────────────────────────────────────────────

  const handleListTeams = React.useCallback(async (): Promise<void> => {
    setState(prev => ({ ...prev, loading: 'ListTeams', error: undefined, teams: [], selectedTeamId: undefined, selectedTeamName: undefined, channels: [] }));
    try {
      const result = await client.callTool('ListTeams', { userId });
      const arr = extractArray(result);
      setState(prev => ({
        ...prev,
        loading: undefined,
        teams: arr.map(t => ({
          displayName: String(t.displayName || ''),
          description: String(t.description || ''),
          id: String(t.id || '')
        }))
      }));
    } catch (err) { setState(prev => ({ ...prev, loading: undefined, error: (err as Error).message })); }
  }, [client, userId]);

  // ── Select Team → List Channels ───────────────────────────

  const handleSelectTeam = React.useCallback(async (teamId: string, teamName: string): Promise<void> => {
    setState(prev => ({ ...prev, selectedTeamId: teamId, selectedTeamName: teamName, channels: [], loading: 'ListChannels', error: undefined }));
    try {
      const result = await client.callTool('ListChannels', { teamId });
      const arr = extractArray(result);
      setState(prev => ({
        ...prev,
        loading: undefined,
        channels: arr.map(c => ({
          displayName: String(c.displayName || ''),
          description: String(c.description || ''),
          id: String(c.id || '')
        }))
      }));
    } catch (err) { setState(prev => ({ ...prev, loading: undefined, error: (err as Error).message })); }
  }, [client]);

  // ── List Chats ────────────────────────────────────────────

  // ── Send Message to Self ──────────────────────────────────

  const handleSendToSelf = React.useCallback(async (): Promise<void> => {
    if (!state.selfMessage.trim()) return;
    setState(prev => ({ ...prev, loading: 'SendMessageToSelf', error: undefined, selfMessageResult: undefined }));
    try {
      const result = await client.callTool('SendMessageToSelf', { content: state.selfMessage.trim() });
      const isError = result.isError || (result.content?.[0]?.text?.indexOf('Error') === 0);
      setState(prev => ({
        ...prev,
        loading: undefined,
        selfMessageResult: isError ? undefined : 'Message sent to yourself!',
        error: isError ? (result.content?.[0]?.text || 'Send failed') : undefined
      }));
    } catch (err) { setState(prev => ({ ...prev, loading: undefined, error: (err as Error).message })); }
  }, [client, state.selfMessage]);

  // ── Render ────────────────────────────────────────────────

  return (
    <div style={{ marginTop: 12 }}>
      <Text variant="small" styles={{ root: { color: theme.palette.neutralTertiary, marginBottom: 12, display: 'block' } }}>
        Explore your Teams, channels, and send a message to yourself.
      </Text>

      {state.error && (
        <Text variant="small" styles={{ root: { color: theme.palette.red, marginBottom: 8, display: 'block' } }}>{state.error}</Text>
      )}

      {/* ── Teams + Channels ─────────────────────────────────── */}
      <div style={cardStyle}>
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }} styles={{ root: { marginBottom: 8 } }}>
          <Text styles={{ root: { fontWeight: 600 } }}>My Teams</Text>
          <Text variant="small" styles={{ root: { color: theme.palette.neutralTertiary } }}>(ListTeams → select → ListChannels)</Text>
        </Stack>
        <PrimaryButton text={state.loading === 'ListTeams' ? 'Loading...' : 'Load My Teams'} iconProps={{ iconName: 'TeamsLogo16' }} onClick={handleListTeams} disabled={isLoading} />

        {state.teams.length > 0 && (
          <div style={{ marginTop: 8 }}>
            <ChoiceGroup
              selectedKey={state.selectedTeamId}
              options={state.teams.map((team): IChoiceGroupOption => ({
                key: team.id,
                text: `${team.displayName}${team.description.trim() ? ' — ' + team.description.trim() : ''}`
              }))}
              onChange={(_, option) => {
                if (option) {
                  const team = state.teams.find(t => t.id === option.key);
                  if (team) { void handleSelectTeam(team.id, team.displayName); }
                }
              }}
            />
          </div>
        )}

        {state.loading === 'ListChannels' && <Spinner size={SpinnerSize.small} label="Loading channels..." styles={{ root: { marginTop: 8 } }} />}

        {state.channels.length > 0 && (
          <div style={{ marginTop: 8, padding: 8, backgroundColor: theme.palette.neutralLighterAlt, borderRadius: 4 }}>
            <Text variant="small" styles={{ root: { fontWeight: 600, color: theme.palette.neutralSecondary, marginBottom: 4, display: 'block' } }}>
              Channels in {state.selectedTeamName} — {state.channels.length} channel{state.channels.length !== 1 ? 's' : ''}
            </Text>
            <Stack tokens={{ childrenGap: 4 }}>
              {state.channels.map((ch, i) => (
                <Stack key={i} horizontal verticalAlign="center" tokens={{ childrenGap: 6 }}>
                  <Icon iconName="NumberSymbol" styles={{ root: { color: '#6264A7', fontSize: 12 } }} />
                  <Text variant="small" styles={{ root: { fontWeight: 600 } }}>{ch.displayName}</Text>
                  {ch.description.trim() && <Text variant="small" styles={{ root: { color: theme.palette.neutralTertiary } }}>— {ch.description.trim()}</Text>}
                </Stack>
              ))}
            </Stack>
          </div>
        )}
      </div>

      {/* ── Send Message to Self ──────────────────────────────── */}
      <div style={cardStyle}>
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }} styles={{ root: { marginBottom: 8 } }}>
          <Text styles={{ root: { fontWeight: 600 } }}>Send Message to Self</Text>
          <Text variant="small" styles={{ root: { color: theme.palette.neutralTertiary } }}>(SendMessageToSelf)</Text>
        </Stack>
        <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="end">
          <TextField
            placeholder="Type a message"
            value={state.selfMessage}
            onChange={(_, val) => setState(prev => ({ ...prev, selfMessage: val || '' }))}
            styles={{ root: { width: 350 } }}
            onKeyDown={(e) => { if (e.key === 'Enter') { void handleSendToSelf(); } }}
          />
          <DefaultButton
            text={state.loading === 'SendMessageToSelf' ? 'Sending...' : 'Send'}
            iconProps={{ iconName: 'Send' }}
            onClick={handleSendToSelf}
            disabled={isLoading || !state.selfMessage.trim()}
          />
        </Stack>
        {state.selfMessageResult && (
          <Text variant="small" styles={{ root: { color: '#107c10', fontWeight: 600, marginTop: 4 } }}>{state.selfMessageResult}</Text>
        )}
      </div>

      {isLoading && state.loading !== 'ListChannels' && <Spinner size={SpinnerSize.small} label={`Calling ${state.loading}...`} styles={{ root: { marginTop: 8 } }} />}
    </div>
  );
};
