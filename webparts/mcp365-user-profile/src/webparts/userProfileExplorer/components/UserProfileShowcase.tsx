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
  Persona,
  PersonaSize,
  getTheme
} from '@fluentui/react';
import type { McpBrowserClient, IMcpCallResult, McpLogHandler } from '../services/McpBrowserClient';

// ─── Types ──────────────────────────────────────────────────────────

interface IShowcaseProps {
  client: McpBrowserClient;
  theme: ReturnType<typeof getTheme>;
  onLog: McpLogHandler;
}

interface IUserCard {
  displayName?: string;
  mail?: string;
  jobTitle?: string;
  department?: string;
  officeLocation?: string;
}

interface IShowcaseState {
  myProfile: IUserCard | undefined;
  myManager: IUserCard | undefined;
  directReports: IUserCard[];
  lookupResult: IUserCard | undefined;
  lookupQuery: string;
  loading: string | undefined;
  error: string | undefined;
}

// ─── Helpers ────────────────────────────────────────────────────────

function extractUserCard(result: IMcpCallResult): IUserCard | undefined {
  if (!result.content || result.content.length === 0) return undefined;
  const text = result.content[0].text;
  if (!text) return undefined;

  const jsonStart = text.indexOf('{');
  if (jsonStart === -1) return undefined;

  try {
    const parsed = JSON.parse(text.substring(jsonStart));
    return {
      displayName: parsed.displayName,
      mail: parsed.mail || parsed.userPrincipalName,
      jobTitle: parsed.jobTitle,
      department: parsed.department,
      officeLocation: parsed.officeLocation
    };
  } catch {
    return undefined;
  }
}

function extractUserCards(result: IMcpCallResult): IUserCard[] {
  if (!result.content || result.content.length === 0) return [];
  const text = result.content[0].text;
  if (!text) return [];

  const jsonStart = text.indexOf('[');
  if (jsonStart !== -1) {
    try {
      const arr = JSON.parse(text.substring(jsonStart)) as Array<Record<string, unknown>>;
      return arr.map(u => ({
        displayName: u.displayName as string,
        mail: (u.mail || u.userPrincipalName) as string,
        jobTitle: u.jobTitle as string,
        department: u.department as string
      }));
    } catch { /* fall through */ }
  }

  // Single object fallback
  const card = extractUserCard(result);
  return card ? [card] : [];
}

// ─── Component ──────────────────────────────────────────────────────

export const UserProfileShowcase: React.FC<IShowcaseProps> = ({ client, theme }) => {
  const [state, setState] = React.useState<IShowcaseState>({
    myProfile: undefined,
    myManager: undefined,
    directReports: [],
    lookupResult: undefined,
    lookupQuery: '',
    loading: undefined,
    error: undefined
  });

  const callTool = React.useCallback(async (toolName: string, args: Record<string, unknown>): Promise<IMcpCallResult> => {
    setState(prev => ({ ...prev, loading: toolName, error: undefined }));
    try {
      const result = await client.callTool(toolName, args);
      setState(prev => ({ ...prev, loading: undefined }));
      return result;
    } catch (err) {
      setState(prev => ({ ...prev, loading: undefined, error: (err as Error).message }));
      throw err;
    }
  }, [client]);

  const handleMyProfile = React.useCallback(async (): Promise<void> => {
    const result = await callTool('GetMyDetails', { select: 'displayName,mail,jobTitle,department,officeLocation' });
    setState(prev => ({ ...prev, myProfile: extractUserCard(result) }));
  }, [callTool]);

  const handleMyManager = React.useCallback(async (): Promise<void> => {
    const result = await callTool('GetManagerDetails', { userId: 'me', select: 'displayName,mail,jobTitle,department' });
    setState(prev => ({ ...prev, myManager: extractUserCard(result) }));
  }, [callTool]);

  const handleDirectReports = React.useCallback(async (): Promise<void> => {
    const result = await callTool('GetDirectReportsDetails', { userId: 'me', select: 'displayName,mail,jobTitle,department' });
    setState(prev => ({ ...prev, directReports: extractUserCards(result) }));
  }, [callTool]);

  const handleLookup = React.useCallback(async (): Promise<void> => {
    if (!state.lookupQuery.trim()) return;
    const result = await callTool('GetUserDetails', { userIdentifier: state.lookupQuery.trim(), select: 'displayName,mail,jobTitle,department,officeLocation' });
    setState(prev => ({ ...prev, lookupResult: extractUserCard(result) }));
  }, [callTool, state.lookupQuery]);

  // ── Render helpers ────────────────────────────────────────────

  const cardStyle: React.CSSProperties = {
    border: `1px solid ${theme.palette.neutralLight}`,
    borderRadius: 8,
    padding: 16,
    backgroundColor: theme.palette.white
  };

  const renderUserCard = (user: IUserCard | undefined, label: string): React.ReactElement | null => {
    if (!user) return null;
    return (
      <div style={cardStyle}>
        <Text variant="small" styles={{ root: { fontWeight: 600, color: theme.palette.neutralSecondary, marginBottom: 8, display: 'block' } }}>
          {label}
        </Text>
        <Persona
          text={user.displayName || 'Unknown'}
          secondaryText={user.jobTitle || ''}
          tertiaryText={user.mail || ''}
          size={PersonaSize.size48}
          showSecondaryText={true}
        />
        {(user.department || user.officeLocation) && (
          <Stack horizontal tokens={{ childrenGap: 16 }} styles={{ root: { marginTop: 8, paddingLeft: 56 } }}>
            {user.department && (
              <Text variant="small" styles={{ root: { color: theme.palette.neutralSecondary } }}>
                <Icon iconName="Org" styles={{ root: { marginRight: 4 } }} />{user.department}
              </Text>
            )}
            {user.officeLocation && (
              <Text variant="small" styles={{ root: { color: theme.palette.neutralSecondary } }}>
                <Icon iconName="POI" styles={{ root: { marginRight: 4 } }} />{user.officeLocation}
              </Text>
            )}
          </Stack>
        )}
      </div>
    );
  };

  const isLoading = !!state.loading;

  return (
    <div style={{ marginTop: 12 }}>
      <Text variant="small" styles={{ root: { color: theme.palette.neutralTertiary, marginBottom: 12, display: 'block' } }}>
        Click the buttons below to call MCP tools and see the results. No JSON, no parameters — just the data.
      </Text>

      {state.error && (
        <Text variant="small" styles={{ root: { color: theme.palette.red, marginBottom: 8, display: 'block' } }}>
          {state.error}
        </Text>
      )}

      {/* Action buttons */}
      <Stack horizontal tokens={{ childrenGap: 8 }} wrap styles={{ root: { marginBottom: 16 } }}>
        <PrimaryButton
          text={state.loading === 'GetMyDetails' ? 'Loading...' : 'My Profile'}
          iconProps={{ iconName: 'Contact' }}
          onClick={handleMyProfile}
          disabled={isLoading}
        />
        <DefaultButton
          text={state.loading === 'GetManagerDetails' ? 'Loading...' : 'My Manager'}
          iconProps={{ iconName: 'People' }}
          onClick={handleMyManager}
          disabled={isLoading}
        />
        <DefaultButton
          text={state.loading === 'GetDirectReportsDetails' ? 'Loading...' : 'My Direct Reports'}
          iconProps={{ iconName: 'Group' }}
          onClick={handleDirectReports}
          disabled={isLoading}
        />
      </Stack>

      {/* Lookup */}
      <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="end" styles={{ root: { marginBottom: 16 } }}>
        <TextField
          label="Lookup user"
          placeholder="UPN or email (e.g. user@contoso.com)"
          value={state.lookupQuery}
          onChange={(_, val) => setState(prev => ({ ...prev, lookupQuery: val || '' }))}
          styles={{ root: { width: 300 } }}
          onKeyDown={(e) => { if (e.key === 'Enter') { void handleLookup(); } }}
        />
        <DefaultButton
          text={state.loading === 'GetUserDetails' ? 'Loading...' : 'Lookup'}
          iconProps={{ iconName: 'Search' }}
          onClick={handleLookup}
          disabled={isLoading || !state.lookupQuery.trim()}
        />
      </Stack>

      {isLoading && <Spinner size={SpinnerSize.small} label={`Calling ${state.loading}...`} styles={{ root: { marginBottom: 12 } }} />}

      {/* Results */}
      <Stack tokens={{ childrenGap: 12 }}>
        {renderUserCard(state.myProfile, 'My Profile (GetMyDetails)')}
        {renderUserCard(state.myManager, 'My Manager (GetManagerDetails)')}
        {state.directReports.length > 0 && (
          <div style={cardStyle}>
            <Text variant="small" styles={{ root: { fontWeight: 600, color: theme.palette.neutralSecondary, marginBottom: 8, display: 'block' } }}>
              My Direct Reports (GetDirectReportsDetails) — {state.directReports.length} people
            </Text>
            <Stack tokens={{ childrenGap: 8 }}>
              {state.directReports.map((user, i) => (
                <Persona
                  key={i}
                  text={user.displayName || 'Unknown'}
                  secondaryText={`${user.jobTitle || ''} ${user.department ? '· ' + user.department : ''}`}
                  tertiaryText={user.mail || ''}
                  size={PersonaSize.size32}
                  showSecondaryText={true}
                />
              ))}
            </Stack>
          </div>
        )}
        {renderUserCard(state.lookupResult, `Lookup Result (GetUserDetails: ${state.lookupQuery})`)}
      </Stack>
    </div>
  );
};
