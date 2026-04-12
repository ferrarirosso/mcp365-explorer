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
  getTheme
} from '@fluentui/react';
import type { McpBrowserClient, IMcpCallResult, McpLogHandler } from '../services/McpBrowserClient';

interface IShowcaseProps {
  client: McpBrowserClient;
  theme: ReturnType<typeof getTheme>;
  onLog: McpLogHandler;
  userEmail: string;
  userId: string;
}

interface IDriveInfo {
  name: string;
  webUrl: string;
  driveType: string;
  ownerDisplayName: string;
  ownerEmail: string;
  fileCount: number | undefined;
  used: number | undefined;
  total: number | undefined;
}

interface ISearchHit {
  id: string;
  name: string;
  webUrl: string;
  size: number | undefined;
  isFolder: boolean;
  lastModified: string | undefined;
}

interface IMetadata {
  name: string;
  webUrl: string;
  size: number | undefined;
  id: string;
  createdDateTime: string | undefined;
  lastModifiedDateTime: string | undefined;
  raw: string;
}

interface IShowcaseState {
  drive: IDriveInfo | undefined;

  searchQuery: string;
  hits: ISearchHit[];

  selectedHitUrl: string | undefined;
  metadata: IMetadata | undefined;

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
    if (outer && typeof outer === 'object' && !Array.isArray(outer) && typeof (outer as Record<string, unknown>).response === 'string') {
      try { return JSON.parse((outer as Record<string, string>).response); } catch { return outer; }
    }
    return outer;
  } catch { return undefined; }
}

function extractObject(result: IMcpCallResult): Record<string, unknown> | undefined {
  const parsed = extractJsonFromContent(result);
  if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) return undefined;
  return parsed as Record<string, unknown>;
}

function extractArray(result: IMcpCallResult): Array<Record<string, unknown>> {
  const parsed = extractJsonFromContent(result);
  if (!parsed) return [];
  if (Array.isArray(parsed)) return parsed as Array<Record<string, unknown>>;
  if (parsed && typeof parsed === 'object') {
    const obj = parsed as Record<string, unknown>;
    if (Array.isArray(obj.value)) return obj.value as Array<Record<string, unknown>>;
    if (Array.isArray(obj.results)) return obj.results as Array<Record<string, unknown>>;
    if (Array.isArray(obj.items)) return obj.items as Array<Record<string, unknown>>;
    return [obj];
  }
  return [];
}

/**
 * Detect error responses. The server sometimes returns isError=false but
 * encodes the failure in the payload as { Error, StatusCode, StatusDescription }.
 * Returns a user-facing error string if present, undefined otherwise.
 */
function extractError(result: IMcpCallResult): string | undefined {
  if (result.isError) {
    return result.content?.[0]?.text || 'Tool call failed';
  }
  const obj = extractObject(result);
  if (obj && typeof obj.Error === 'string') {
    const code = typeof obj.StatusCode === 'number' ? ` (${obj.StatusCode}${typeof obj.StatusDescription === 'string' ? ' ' + obj.StatusDescription : ''})` : '';
    return `${obj.Error}${code}`;
  }
  return undefined;
}

function getString(obj: Record<string, unknown> | undefined, key: string): string | undefined {
  if (!obj) return undefined;
  const v = obj[key];
  return typeof v === 'string' ? v : undefined;
}

function getNumber(obj: Record<string, unknown> | undefined, key: string): number | undefined {
  if (!obj) return undefined;
  const v = obj[key];
  return typeof v === 'number' ? v : undefined;
}

function formatBytes(bytes: number | undefined): string {
  if (bytes === undefined || bytes === null) return '—';
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  if (bytes < 1024 * 1024 * 1024) return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  return `${(bytes / (1024 * 1024 * 1024)).toFixed(2)} GB`;
}

export const OneDriveShowcase: React.FC<IShowcaseProps> = ({ client, theme }) => {
  const [state, setState] = React.useState<IShowcaseState>({
    drive: undefined,
    searchQuery: '',
    hits: [],
    selectedHitUrl: undefined,
    metadata: undefined,
    loading: undefined,
    error: undefined
  });

  const isLoading = !!state.loading;

  const cardStyle: React.CSSProperties = {
    border: `1px solid ${theme.palette.neutralLight}`, borderRadius: 8, padding: 12, backgroundColor: theme.palette.white, marginTop: 8
  };

  const stepBadgeStyle: React.CSSProperties = {
    display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
    width: 22, height: 22, borderRadius: '50%',
    backgroundColor: theme.palette.themePrimary, color: theme.palette.white,
    fontSize: 12, fontWeight: 600, marginRight: 8
  };

  // ── Step 1: getOnedrive ────────────────────────────────────

  const handleGetOnedrive = React.useCallback(async (): Promise<void> => {
    setState(prev => ({ ...prev, loading: 'getOnedrive', error: undefined, drive: undefined }));
    try {
      const result = await client.callTool('getOnedrive', {});
      const err = extractError(result);
      if (err) {
        setState(prev => ({ ...prev, loading: undefined, error: err }));
        return;
      }
      const obj = extractObject(result);
      const ownerUser = obj && typeof obj.owner === 'object' ? (obj.owner as Record<string, unknown>).user as Record<string, unknown> | undefined : undefined;
      const quota = obj && typeof obj.quota === 'object' ? obj.quota as Record<string, unknown> : undefined;
      const drive: IDriveInfo = {
        name: getString(obj, 'name') || 'OneDrive',
        webUrl: getString(obj, 'webUrl') || '',
        driveType: getString(obj, 'driveType') || '',
        ownerDisplayName: getString(ownerUser, 'displayName') || '',
        ownerEmail: getString(ownerUser, 'email') || '',
        fileCount: getNumber(quota, 'fileCount'),
        used: getNumber(quota, 'used'),
        total: getNumber(quota, 'total')
      };
      setState(prev => ({ ...prev, loading: undefined, drive }));
    } catch (err) { setState(prev => ({ ...prev, loading: undefined, error: (err as Error).message })); }
  }, [client]);

  // ── Step 2: findFileOrFolderInMyDrive ─────────────────────

  const handleSearch = React.useCallback(async (): Promise<void> => {
    if (!state.searchQuery.trim()) return;
    setState(prev => ({ ...prev, loading: 'findFileOrFolderInMyDrive', error: undefined, hits: [], metadata: undefined, selectedHitUrl: undefined }));
    try {
      const result = await client.callTool('findFileOrFolderInMyDrive', { searchQuery: state.searchQuery.trim() });
      const err = extractError(result);
      if (err) {
        setState(prev => ({ ...prev, loading: undefined, error: err }));
        return;
      }
      const arr = extractArray(result);
      const hits: ISearchHit[] = arr.map(item => ({
        id: getString(item, 'id') || '',
        name: getString(item, 'name') || '(unnamed)',
        webUrl: getString(item, 'webUrl') || '',
        size: getNumber(item, 'size'),
        isFolder: !!item.folder,
        lastModified: getString(item, 'lastModifiedDateTime')
      }));
      setState(prev => ({ ...prev, loading: undefined, hits }));
    } catch (err) { setState(prev => ({ ...prev, loading: undefined, error: (err as Error).message })); }
  }, [client, state.searchQuery]);

  // ── Step 3: getFileOrFolderMetadataInMyOnedrive ────────────
  //
  // We use the ID-based variant (not ByUrl) because search hits return
  // DispForm.aspx webUrls which the URL-based metadata call can't resolve.
  // The ID from findFileOrFolderInMyDrive is stable and works directly.

  const handleGetMetadata = React.useCallback(async (hit: ISearchHit): Promise<void> => {
    if (!hit.id) return;
    setState(prev => ({ ...prev, loading: 'getFileOrFolderMetadataInMyOnedrive', error: undefined, metadata: undefined, selectedHitUrl: hit.webUrl }));
    try {
      const result = await client.callTool('getFileOrFolderMetadataInMyOnedrive', { fileOrFolderId: hit.id });
      const err = extractError(result);
      if (err) {
        setState(prev => ({ ...prev, loading: undefined, error: err }));
        return;
      }
      const obj = extractObject(result);
      const raw = result.content?.[0]?.text || '';
      const metadata: IMetadata = {
        name: getString(obj, 'name') || hit.name,
        webUrl: getString(obj, 'webUrl') || hit.webUrl,
        size: getNumber(obj, 'size'),
        id: getString(obj, 'id') || hit.id,
        createdDateTime: getString(obj, 'createdDateTime'),
        lastModifiedDateTime: getString(obj, 'lastModifiedDateTime'),
        raw
      };
      setState(prev => ({ ...prev, loading: undefined, metadata }));
    } catch (err) { setState(prev => ({ ...prev, loading: undefined, error: (err as Error).message })); }
  }, [client]);

  // ── Render ────────────────────────────────────────────────

  const quotaPct = state.drive && state.drive.used !== undefined && state.drive.total
    ? Math.min(100, Math.round((state.drive.used / state.drive.total) * 100))
    : undefined;

  return (
    <div style={{ marginTop: 12 }}>
      <Text variant="small" styles={{ root: { color: theme.palette.neutralTertiary, marginBottom: 12, display: 'block' } }}>
        Explore your OneDrive — overview, search, and item metadata in three chained MCP calls.
      </Text>

      {state.error && (
        <Text variant="small" styles={{ root: { color: theme.palette.red, marginBottom: 8, display: 'block' } }}>{state.error}</Text>
      )}

      {/* ── Step 1: Overview ───────────────────────────────────── */}
      <div style={cardStyle}>
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }} styles={{ root: { marginBottom: 8 } }}>
          <span style={stepBadgeStyle}>1</span>
          <Text styles={{ root: { fontWeight: 600 } }}>My OneDrive</Text>
          <Text variant="small" styles={{ root: { color: theme.palette.neutralTertiary } }}>(getOnedrive)</Text>
        </Stack>
        <PrimaryButton
          text={state.loading === 'getOnedrive' ? 'Loading...' : 'Show my OneDrive'}
          iconProps={{ iconName: 'OneDrive' }}
          onClick={handleGetOnedrive}
          disabled={isLoading}
        />

        {state.drive && (
          <div style={{ marginTop: 10, padding: 10, backgroundColor: theme.palette.neutralLighterAlt, borderRadius: 4 }}>
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
              <Icon iconName="OneDrive" styles={{ root: { color: theme.palette.themePrimary, fontSize: 20 } }} />
              <Text styles={{ root: { fontWeight: 600 } }}>{state.drive.name}</Text>
              {state.drive.driveType && (
                <Text variant="small" styles={{ root: { color: theme.palette.neutralTertiary, textTransform: 'capitalize' } }}>
                  · {state.drive.driveType}
                </Text>
              )}
            </Stack>
            {state.drive.ownerDisplayName && (
              <Text variant="small" styles={{ root: { display: 'block', marginTop: 4 } }}>
                Owner: <strong>{state.drive.ownerDisplayName}</strong>
                {state.drive.ownerEmail && ` · ${state.drive.ownerEmail}`}
              </Text>
            )}
            {state.drive.webUrl && (
              <Text variant="small" styles={{ root: { display: 'block', marginTop: 2, wordBreak: 'break-all' } }}>
                <a href={state.drive.webUrl} target="_blank" rel="noreferrer">{state.drive.webUrl}</a>
              </Text>
            )}
            {quotaPct !== undefined && (
              <div style={{ marginTop: 8 }}>
                <Text variant="small" styles={{ root: { display: 'block', marginBottom: 4 } }}>
                  Storage: <strong>{formatBytes(state.drive.used)}</strong> of {formatBytes(state.drive.total)} ({quotaPct}%)
                  {state.drive.fileCount !== undefined && ` · ${state.drive.fileCount} file${state.drive.fileCount !== 1 ? 's' : ''}`}
                </Text>
                <div style={{ height: 6, backgroundColor: theme.palette.neutralLight, borderRadius: 3, overflow: 'hidden' }}>
                  <div style={{ height: '100%', width: `${quotaPct}%`, backgroundColor: theme.palette.themePrimary }} />
                </div>
              </div>
            )}
          </div>
        )}
      </div>

      {/* ── Step 2: Search ─────────────────────────────────────── */}
      <div style={cardStyle}>
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }} styles={{ root: { marginBottom: 8 } }}>
          <span style={stepBadgeStyle}>2</span>
          <Text styles={{ root: { fontWeight: 600 } }}>Find a file or folder</Text>
          <Text variant="small" styles={{ root: { color: theme.palette.neutralTertiary } }}>(findFileOrFolderInMyDrive)</Text>
        </Stack>
        <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="end">
          <TextField
            placeholder="e.g. report, budget, contract"
            value={state.searchQuery}
            onChange={(_, val) => setState(prev => ({ ...prev, searchQuery: val || '' }))}
            disabled={isLoading}
            styles={{ root: { width: 320 } }}
            onKeyDown={(e) => { if (e.key === 'Enter') { void handleSearch(); } }}
          />
          <DefaultButton
            text={state.loading === 'findFileOrFolderInMyDrive' ? 'Searching...' : 'Search'}
            iconProps={{ iconName: 'Search' }}
            onClick={handleSearch}
            disabled={isLoading || !state.searchQuery.trim()}
          />
        </Stack>

        {state.hits.length > 0 && (
          <div style={{ marginTop: 10, padding: 10, backgroundColor: theme.palette.neutralLighterAlt, borderRadius: 4 }}>
            <Text variant="small" styles={{ root: { fontWeight: 600, color: theme.palette.neutralSecondary, marginBottom: 6, display: 'block' } }}>
              {state.hits.length} result{state.hits.length !== 1 ? 's' : ''}
            </Text>
            <Stack tokens={{ childrenGap: 6 }}>
              {state.hits.slice(0, 20).map((hit, i) => (
                <div key={i} style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                  <Icon
                    iconName={hit.isFolder ? 'FabricFolder' : 'Page'}
                    styles={{ root: { color: hit.isFolder ? theme.palette.themePrimary : theme.palette.neutralSecondary, fontSize: 14, flexShrink: 0 } }}
                  />
                  <Text variant="small" styles={{ root: { fontWeight: 600, flexShrink: 0 } }}>{hit.name}</Text>
                  {hit.size !== undefined && !hit.isFolder && (
                    <Text variant="small" styles={{ root: { color: theme.palette.neutralTertiary, flexShrink: 0 } }}>· {formatBytes(hit.size)}</Text>
                  )}
                  <div style={{ flexGrow: 1 }} />
                  {hit.id && (
                    <DefaultButton
                      text="Metadata"
                      iconProps={{ iconName: 'Info' }}
                      onClick={() => handleGetMetadata(hit)}
                      disabled={isLoading}
                      styles={{ root: { minWidth: 100, height: 26, padding: '0 10px', flexShrink: 0 } }}
                    />
                  )}
                </div>
              ))}
              {state.hits.length > 20 && (
                <Text variant="small" styles={{ root: { color: theme.palette.neutralTertiary, fontStyle: 'italic' } }}>
                  ...and {state.hits.length - 20} more
                </Text>
              )}
            </Stack>
          </div>
        )}
      </div>

      {/* ── Step 3: Metadata ───────────────────────────────────── */}
      <div style={cardStyle}>
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }} styles={{ root: { marginBottom: 8 } }}>
          <span style={stepBadgeStyle}>3</span>
          <Text styles={{ root: { fontWeight: 600 } }}>Get item metadata</Text>
          <Text variant="small" styles={{ root: { color: theme.palette.neutralTertiary } }}>(getFileOrFolderMetadataInMyOnedrive — click a result above)</Text>
        </Stack>

        {!state.metadata && !state.selectedHitUrl && (
          <Text variant="small" styles={{ root: { color: theme.palette.neutralTertiary, display: 'block' } }}>
            Pick a result from step 2 and click <strong>Metadata</strong> to fetch its details via <code>getFileOrFolderMetadataInMyOnedrive</code>.
          </Text>
        )}

        {state.metadata && (
          <div style={{ padding: 10, backgroundColor: theme.palette.neutralLighterAlt, borderRadius: 4 }}>
            <Text styles={{ root: { fontWeight: 600, display: 'block' } }}>{state.metadata.name}</Text>
            {state.metadata.size !== undefined && (
              <Text variant="small" styles={{ root: { display: 'block', color: theme.palette.neutralSecondary } }}>
                Size: {formatBytes(state.metadata.size)}
              </Text>
            )}
            {state.metadata.createdDateTime && (
              <Text variant="small" styles={{ root: { display: 'block', color: theme.palette.neutralSecondary } }}>
                Created: {new Date(state.metadata.createdDateTime).toLocaleString()}
              </Text>
            )}
            {state.metadata.lastModifiedDateTime && (
              <Text variant="small" styles={{ root: { display: 'block', color: theme.palette.neutralSecondary } }}>
                Modified: {new Date(state.metadata.lastModifiedDateTime).toLocaleString()}
              </Text>
            )}
            {state.metadata.id && (
              <Text variant="small" styles={{ root: { display: 'block', color: theme.palette.neutralTertiary, fontFamily: 'Consolas, monospace', marginTop: 4, wordBreak: 'break-all' } }}>
                id: {state.metadata.id}
              </Text>
            )}
          </div>
        )}
      </div>

      {isLoading && <Spinner size={SpinnerSize.small} label={`Calling ${state.loading}...`} styles={{ root: { marginTop: 8 } }} />}
    </div>
  );
};
