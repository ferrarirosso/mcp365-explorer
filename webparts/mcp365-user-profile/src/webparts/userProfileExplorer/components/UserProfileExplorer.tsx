import * as React from 'react';
import {
  PrimaryButton,
  DefaultButton,
  TextField,
  Dropdown,
  Stack,
  Text,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  Pivot,
  PivotItem,
  Icon,
  Label,
  getTheme,
  TooltipHost,
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  SearchBox
} from '@fluentui/react';
import { DropdownMenuItemType } from '@fluentui/react';
import type { IDropdownOption, IColumn } from '@fluentui/react';
import type { IUserProfileExplorerProps } from './IUserProfileExplorerProps';
import {
  McpBrowserClient,
  MCP_AUDIENCE
} from '../services/McpBrowserClient';
import type {
  IMcpTool,
  IMcpPrompt,
  IMcpResource,
  IMcpLogEntry,
  IMcpServerInfo,
  IMcpCapabilities
} from '../services/McpBrowserClient';
import {
  ME_SERVER_ID,
  ME_SERVER_EXAMPLES
} from '../services/McpServerCatalog';
import type { IToolExample } from '../services/McpServerCatalog';
import { UserProfileShowcase } from './UserProfileShowcase';

// ─── Preset Storage ─────────────────────────────────────────────────

const PRESETS_STORAGE_KEY = 'mcp365-user-profile-presets';

function loadCustomPresets(): Record<string, IToolExample[]> {
  try {
    const raw = localStorage.getItem(PRESETS_STORAGE_KEY);
    return raw ? JSON.parse(raw) as Record<string, IToolExample[]> : {};
  } catch {
    return {};
  }
}

function saveCustomPresets(presets: Record<string, IToolExample[]>): void {
  localStorage.setItem(PRESETS_STORAGE_KEY, JSON.stringify(presets));
}

// ─── State ──────────────────────────────────────────────────────────

interface IExplorerState {
  connectionState: 'disconnected' | 'connecting' | 'connected' | 'error';
  connectionError: string | undefined;

  // Server info from initialize
  protocolVersion: string | undefined;
  serverInfo: IMcpServerInfo | undefined;
  capabilities: IMcpCapabilities | undefined;
  sessionId: string | undefined;

  // Discovered capabilities
  discoveredTools: IMcpTool[];
  discoveredPrompts: IMcpPrompt[];
  discoveredResources: IMcpResource[];

  // Tool execution
  selectedToolName: string | undefined;
  parameterValues: Record<string, string>;
  isExecuting: boolean;
  lastRequest: unknown;
  lastResponse: unknown;
  lastDurationMs: number | undefined;

  // Navigation
  activeTab: string;

  // Presets
  customPresets: Record<string, IToolExample[]>;
  presetSaveName: string;
  showPresetSave: boolean;

  // Logs
  logs: IMcpLogEntry[];
  showLogs: boolean;
  logFilter: string;
  logExpandedIndex: number | undefined;
}

// ─── Log Viewer Sub-Component ───────────────────────────────────────

interface ILogViewerProps {
  logs: IMcpLogEntry[];
  filter: string;
  expandedIndex: number | undefined;
  onFilterChange: (val: string) => void;
  onToggleExpand: (idx: number) => void;
  theme: ReturnType<typeof getTheme>;
}

function formatTime(d: Date): string {
  return d.toLocaleTimeString('en-US', { hour12: false, hour: '2-digit', minute: '2-digit', second: '2-digit', fractionalSecondDigits: 3 } as Intl.DateTimeFormatOptions);
}

const LogViewer: React.FC<ILogViewerProps> = ({ logs, filter, expandedIndex, onFilterChange, onToggleExpand, theme }) => {
  const [sortKey, setSortKey] = React.useState<string>('');
  const [sortDesc, setSortDesc] = React.useState<boolean>(false);

  const filterLower = filter.toLowerCase();
  let filteredLogs = filterLower
    ? logs.filter((entry) => {
        return entry.message.toLowerCase().indexOf(filterLower) !== -1 ||
               entry.category.toLowerCase().indexOf(filterLower) !== -1 ||
               entry.method.toLowerCase().indexOf(filterLower) !== -1 ||
               entry.level.toLowerCase().indexOf(filterLower) !== -1 ||
               (entry.data && JSON.stringify(entry.data).toLowerCase().indexOf(filterLower) !== -1);
      })
    : logs;

  // Apply sorting
  if (sortKey) {
    filteredLogs = [...filteredLogs].sort((a, b) => {
      let aVal: string | number = '';
      let bVal: string | number = '';
      if (sortKey === 'time') { aVal = a.timestamp.getTime(); bVal = b.timestamp.getTime(); }
      else if (sortKey === 'level') { aVal = a.level; bVal = b.level; }
      else if (sortKey === 'category') { aVal = a.category; bVal = b.category; }
      else if (sortKey === 'method') { aVal = a.method; bVal = b.method; }
      else if (sortKey === 'message') { aVal = a.message; bVal = b.message; }
      else if (sortKey === 'dir') { aVal = a.direction; bVal = b.direction; }

      if (aVal < bVal) return sortDesc ? 1 : -1;
      if (aVal > bVal) return sortDesc ? -1 : 1;
      return 0;
    });
  }

  const handleColumnClick = React.useCallback((_: unknown, column?: IColumn): void => {
    if (!column) return;
    if (sortKey === column.key) {
      setSortDesc(!sortDesc);
    } else {
      setSortKey(column.key);
      setSortDesc(false);
    }
  }, [sortKey, sortDesc]);

  const levelColor = (level: string): string => {
    if (level === 'error') return theme.palette.red;
    if (level === 'warning') return theme.palette.yellowDark;
    if (level === 'info') return theme.palette.themePrimary;
    return theme.palette.neutralSecondary;
  };

  const directionIcon = (dir: string): string => {
    if (dir === 'send') return '\u2192';   // →
    if (dir === 'receive') return '\u2190'; // ←
    return '\u00B7';                        // ·
  };

  const columns: IColumn[] = [
    {
      key: 'time',
      name: 'Time',
      fieldName: 'time',
      minWidth: 85,
      maxWidth: 85,
      isSorted: sortKey === 'time',
      isSortedDescending: sortKey === 'time' && sortDesc,
      onRender: (item: IMcpLogEntry) => (
        <Text variant="small" styles={{ root: { fontFamily: 'Consolas, monospace', fontSize: 11, color: theme.palette.neutralTertiary } }}>
          {formatTime(item.timestamp)}
        </Text>
      )
    },
    {
      key: 'level',
      name: 'Level',
      fieldName: 'level',
      minWidth: 55,
      maxWidth: 55,
      isSorted: sortKey === 'level',
      isSortedDescending: sortKey === 'level' && sortDesc,
      onRender: (item: IMcpLogEntry) => (
        <Text variant="small" styles={{ root: { fontWeight: 600, fontSize: 11, color: levelColor(item.level) } }}>
          {item.level}
        </Text>
      )
    },
    {
      key: 'dir',
      name: '',
      fieldName: 'direction',
      minWidth: 20,
      maxWidth: 20,
      isSorted: sortKey === 'dir',
      isSortedDescending: sortKey === 'dir' && sortDesc,
      onRender: (item: IMcpLogEntry) => (
        <Text variant="small" styles={{ root: { fontSize: 14, fontWeight: 600, color: item.direction === 'send' ? theme.palette.blue : item.direction === 'receive' ? theme.palette.green : theme.palette.neutralTertiary } }}>
          {directionIcon(item.direction)}
        </Text>
      )
    },
    {
      key: 'category',
      name: 'Type',
      fieldName: 'category',
      minWidth: 40,
      maxWidth: 50,
      isSorted: sortKey === 'category',
      isSortedDescending: sortKey === 'category' && sortDesc,
      onRender: (item: IMcpLogEntry) => (
        <Text variant="small" styles={{ root: { fontFamily: 'Consolas, monospace', fontSize: 11, fontWeight: 600 } }}>
          {item.category}
        </Text>
      )
    },
    {
      key: 'method',
      name: 'Method',
      fieldName: 'method',
      minWidth: 100,
      maxWidth: 150,
      isSorted: sortKey === 'method',
      isSortedDescending: sortKey === 'method' && sortDesc,
      onRender: (item: IMcpLogEntry) => (
        <Text variant="small" styles={{ root: { fontFamily: 'Consolas, monospace', fontSize: 11, color: theme.palette.neutralPrimary } }}>
          {item.method || '-'}
        </Text>
      )
    },
    {
      key: 'message',
      name: 'Message',
      fieldName: 'message',
      minWidth: 200,
      isMultiline: false,
      isSorted: sortKey === 'message',
      isSortedDescending: sortKey === 'message' && sortDesc,
      onRender: (item: IMcpLogEntry) => (
        <Text variant="small" styles={{ root: { fontSize: 11 } }}>
          {item.message}
        </Text>
      )
    }
  ];

  return (
    <div style={{ marginTop: 8 }}>
      <SearchBox
        placeholder="Search logs (message, method, type...)"
        value={filter}
        onChange={(_, val) => onFilterChange(val || '')}
        styles={{ root: { marginBottom: 8, maxWidth: 400 } }}
      />
      {filteredLogs.length === 0 ? (
        <Text variant="small" styles={{ root: { color: theme.palette.neutralTertiary, padding: 8 } }}>
          {logs.length === 0 ? 'No log entries yet. Connect to start.' : 'No matches.'}
        </Text>
      ) : (
        <div style={{ maxHeight: 300, overflow: 'auto' }}>
          <DetailsList
            items={filteredLogs}
            columns={columns}
            layoutMode={DetailsListLayoutMode.fixedColumns}
            selectionMode={SelectionMode.none}
            compact={true}
            isHeaderVisible={true}
            onColumnHeaderClick={handleColumnClick}
            onItemInvoked={(item: IMcpLogEntry) => {
              const idx = logs.indexOf(item);
              if (idx !== -1) onToggleExpand(idx);
            }}
            onActiveItemChanged={(item?: IMcpLogEntry) => {
              if (item) {
                const idx = logs.indexOf(item);
                if (idx !== -1) onToggleExpand(idx);
              }
            }}
            onRenderRow={(rowProps, defaultRender) => {
              if (!rowProps || !defaultRender) return null;
              const idx = logs.indexOf(rowProps.item as IMcpLogEntry);
              const isExpanded = idx === expandedIndex;
              const entry = rowProps.item as IMcpLogEntry;
              return (
                <>
                  {defaultRender(rowProps)}
                  {isExpanded && entry.data && (
                    <div style={{
                      padding: '4px 12px 8px 12px',
                      fontFamily: 'Consolas, monospace',
                      fontSize: 11,
                      backgroundColor: theme.palette.neutralLighterAlt,
                      borderBottom: `1px solid ${theme.palette.neutralLight}`,
                      whiteSpace: 'pre-wrap',
                      wordBreak: 'break-word',
                      maxHeight: 200,
                      overflow: 'auto'
                    }}>
                      {typeof entry.data === 'string' ? entry.data : JSON.stringify(entry.data, null, 2)}
                    </div>
                  )}
                </>
              );
            }}
            styles={{
              root: { fontSize: 11 },
              headerWrapper: {
                selectors: {
                  '.ms-DetailsHeader': {
                    paddingTop: 0,
                    height: 28,
                    lineHeight: 28
                  }
                }
              }
            }}
          />
        </div>
      )}
      <Stack horizontal tokens={{ childrenGap: 12 }} verticalAlign="center" styles={{ root: { marginTop: 4 } }}>
        {filteredLogs.length > 0 && (
          <Text variant="small" styles={{ root: { color: theme.palette.neutralTertiary, fontStyle: 'italic' } }}>
            Double-click a row to expand its data payload
          </Text>
        )}
        {filteredLogs.length > 0 && (
          <DefaultButton
            text="Copy logs"
            iconProps={{ iconName: 'Copy' }}
            onClick={() => {
              const text = filteredLogs.map(e => {
                const time = formatTime(e.timestamp);
                const dir = e.direction === 'send' ? '→' : e.direction === 'receive' ? '←' : '·';
                const data = e.data ? '\n  ' + (typeof e.data === 'string' ? e.data : JSON.stringify(e.data, null, 2).replace(/\n/g, '\n  ')) : '';
                return `${time} ${e.level} ${dir} [${e.category}] ${e.method} ${e.message}${data}`;
              }).join('\n');
              void navigator.clipboard.writeText(text);
            }}
            styles={{ root: { minWidth: 0, padding: '0 8px', height: 24 } }}
          />
        )}
        {filter && filteredLogs.length < logs.length && (
          <Text variant="small" styles={{ root: { color: theme.palette.neutralTertiary } }}>
            Showing {filteredLogs.length} of {logs.length} entries
          </Text>
        )}
      </Stack>
    </div>
  );
};

// ─── Formatted Response Sub-Component ────────────────────────────────

/**
 * Strips Graph serialization noise from MCP response objects.
 * Removes backingStore, odataType, empty additionalData, and cleans HTML body content.
 */
function cleanGraphObject(obj: unknown): unknown {
  if (obj === null || obj === undefined) return obj;
  if (typeof obj !== 'object') return obj;

  if (Array.isArray(obj)) {
    return obj.map(cleanGraphObject);
  }

  const input = obj as Record<string, unknown>;
  const result: Record<string, unknown> = {};

  Object.keys(input).forEach(key => {
    // Skip serialization noise
    if (key === 'backingStore' || key === 'odataType') return;

    // Skip additionalData if it only has @odata keys or is empty
    if (key === 'additionalData') {
      const ad = input[key] as Record<string, unknown> | undefined;
      if (!ad || typeof ad !== 'object') return;
      const adKeys = Object.keys(ad);
      const hasOnlyOdata = adKeys.length === 0 || adKeys.every(k => k.indexOf('@odata') === 0);
      if (hasOnlyOdata) return;
      // Keep non-odata additional data
      const cleaned: Record<string, unknown> = {};
      adKeys.forEach(k => { if (k.indexOf('@odata') !== 0) cleaned[k] = ad[k]; });
      if (Object.keys(cleaned).length > 0) result[key] = cleaned;
      return;
    }

    // For body.content with HTML, extract plain text preview instead
    if (key === 'content' && typeof input[key] === 'string') {
      const content = input[key] as string;
      if (content.indexOf('<html') !== -1 || content.indexOf('<div') !== -1) {
        // Strip HTML tags for a clean preview
        const plain = content
          .replace(/<br\s*\/?>/gi, '\n')
          .replace(/<\/div>/gi, '\n')
          .replace(/<\/p>/gi, '\n')
          .replace(/<[^>]+>/g, '')
          .replace(/&nbsp;/g, ' ')
          .replace(/&amp;/g, '&')
          .replace(/\r\n/g, '\n')
          .replace(/\n{3,}/g, '\n\n')
          .trim();
        // Keep first meaningful part (before Teams boilerplate)
        const teamsIdx = plain.indexOf('________________');
        result[key] = teamsIdx > 0 ? plain.substring(0, teamsIdx).trim() : plain;
        return;
      }
    }

    // Skip contentType numeric enum (not useful to developers)
    if (key === 'contentType' && typeof input[key] === 'number') return;

    result[key] = cleanGraphObject(input[key]);
  });

  return result;
}

/**
 * Tries to parse embedded JSON from MCP content text blocks.
 * MCP responses often contain a message prefix followed by a JSON string.
 */
function tryParseEmbeddedJson(text: string): { message: string; parsed: unknown } | undefined {
  // Look for JSON object or array starting with { or [
  const jsonStart = text.indexOf('{');
  const arrStart = text.indexOf('[');
  const start = jsonStart === -1 ? arrStart : (arrStart === -1 ? jsonStart : Math.min(jsonStart, arrStart));

  if (start === -1) return undefined;

  const message = text.substring(0, start).trim();
  const jsonPart = text.substring(start);

  try {
    const parsed = JSON.parse(jsonPart);
    return { message, parsed };
  } catch {
    return undefined;
  }
}

interface IFormattedResponseProps {
  response: unknown;
  theme: ReturnType<typeof getTheme>;
}

const FormattedResponse: React.FC<IFormattedResponseProps> = ({ response, theme }) => {
  const resp = response as { content?: Array<{ type: string; text?: string }>; isError?: boolean; error?: string };

  // Error response
  if (resp.error) {
    return (
      <div style={{ padding: 12, color: theme.palette.red, fontWeight: 600 }}>
        {resp.error}
      </div>
    );
  }

  // No content array — render as raw JSON
  if (!resp.content || !Array.isArray(resp.content)) {
    return (
      <div style={{ padding: 12, fontFamily: 'Consolas, monospace', fontSize: 12, whiteSpace: 'pre-wrap' }}>
        {JSON.stringify(response, null, 2)}
      </div>
    );
  }

  const codeStyle: React.CSSProperties = {
    fontFamily: 'Consolas, Monaco, "Courier New", monospace',
    fontSize: 12,
    backgroundColor: theme.palette.neutralLighterAlt,
    padding: 10,
    borderRadius: 4,
    overflow: 'auto',
    maxHeight: 350,
    whiteSpace: 'pre-wrap',
    wordBreak: 'break-word'
  };

  return (
    <Stack tokens={{ childrenGap: 8 }} styles={{ root: { padding: 12 } }}>
      {resp.isError && (
        <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
          The server returned an error response
        </MessageBar>
      )}
      {resp.content.map((block, i) => {
        if (block.type !== 'text' || !block.text) {
          return (
            <div key={i} style={codeStyle}>
              {JSON.stringify(block, null, 2)}
            </div>
          );
        }

        const embedded = tryParseEmbeddedJson(block.text);

        if (embedded) {
          return (
            <div key={i}>
              {embedded.message && (
                <Text variant="small" styles={{ root: { fontWeight: 600, marginBottom: 4, display: 'block', color: theme.palette.neutralPrimary } }}>
                  {embedded.message}
                </Text>
              )}
              <div style={codeStyle}>
                {JSON.stringify(cleanGraphObject(embedded.parsed), null, 2)}
              </div>
            </div>
          );
        }

        // Plain text — might be a correlation ID or metadata
        const isMetadata = block.text.indexOf('CorrelationId:') !== -1 || block.text.indexOf('TimeStamp:') !== -1;
        return (
          <Text key={i} variant="small" styles={{
            root: {
              color: isMetadata ? theme.palette.neutralTertiary : theme.palette.neutralPrimary,
              fontFamily: isMetadata ? 'Consolas, monospace' : undefined,
              fontSize: isMetadata ? 11 : 13
            }
          }}>
            {block.text}
          </Text>
        );
      })}
    </Stack>
  );
};

// ─── Main Component ─────────────────────────────────────────────────

const INITIAL_STATE: IExplorerState = {
  connectionState: 'disconnected',
  connectionError: undefined,
  protocolVersion: undefined,
  serverInfo: undefined,
  capabilities: undefined,
  sessionId: undefined,
  discoveredTools: [],
  discoveredPrompts: [],
  discoveredResources: [],
  selectedToolName: undefined,
  parameterValues: {},
  isExecuting: false,
  lastRequest: undefined,
  lastResponse: undefined,
  lastDurationMs: undefined,
  activeTab: 'tools',
  customPresets: loadCustomPresets(),
  presetSaveName: '',
  showPresetSave: false,
  logs: [],
  showLogs: true,
  logFilter: '',
  logExpandedIndex: undefined
};

// ─── Component ──────────────────────────────────────────────────────

export const UserProfileExplorer: React.FC<IUserProfileExplorerProps> = (props) => {
  const [state, setState] = React.useState<IExplorerState>(INITIAL_STATE);
  const clientRef = React.useRef<McpBrowserClient | undefined>(undefined);
  const theme = getTheme();

  // ── Helpers ──────────────────────────────────────────────────────

  const addLog = React.useCallback((entry: IMcpLogEntry): void => {
    setState(prev => ({ ...prev, logs: [...prev.logs, entry] }));
  }, []);

  const getToken = React.useCallback(async (): Promise<string> => {
    if (!props.tokenProvider) {
      throw new Error('Token provider not available');
    }
    return props.tokenProvider.getToken(MCP_AUDIENCE);
  }, [props.tokenProvider]);

  // ── Derived values ──────────────────────────────────────────────

  const toolList: IMcpTool[] = state.discoveredTools;

  const selectedTool = toolList.find(t => t.name === state.selectedToolName);
  const selectedSchema = selectedTool?.inputSchema as { type: string; properties: Record<string, { type: string; description: string }>; required?: string[] } | undefined;
  const selectedExamples: IToolExample[] = state.selectedToolName
    ? (ME_SERVER_EXAMPLES[state.selectedToolName] || [])
    : [];

  // ── Connect / Disconnect ────────────────────────────────────────

  const handleConnect = React.useCallback(async (): Promise<void> => {
    if (!props.environmentId) {
      setState(prev => ({ ...prev, connectionError: 'Environment ID is required. Configure it in the web part property pane.' }));
      return;
    }

    setState(prev => ({ ...prev, connectionState: 'connecting', connectionError: undefined }));

    try {
      const client = new McpBrowserClient({
        environmentId: props.environmentId,
        serverId: ME_SERVER_ID,
        getToken,
        onLog: addLog
      });

      const result = await client.connect();
      clientRef.current = client;

      setState(prev => ({
        ...prev,
        connectionState: 'connected',
        activeTab: 'showcase',
        protocolVersion: result.protocolVersion,
        serverInfo: result.serverInfo,
        capabilities: result.capabilities,
        sessionId: result.sessionId,
        discoveredTools: result.tools,
        discoveredPrompts: result.prompts,
        discoveredResources: result.resources,
        selectedToolName: result.tools.length > 0 ? result.tools[0].name : undefined,
        parameterValues: {}
      }));
    } catch (err) {
      const message = (err as Error).message || 'Unknown error';
      const isCors = message.indexOf('Failed to fetch') !== -1 ||
                     message.indexOf('CORS') !== -1 ||
                     message.indexOf('NetworkError') !== -1;

      setState(prev => ({
        ...prev,
        connectionState: 'error',
        connectionError: isCors
          ? `CORS blocked: The Agent 365 gateway does not allow direct browser requests. (${message})`
          : message
      }));
    }
  }, [props.environmentId, getToken, addLog]);

  const handleDisconnect = React.useCallback((): void => {
    clientRef.current = undefined;
    setState(prev => ({
      ...INITIAL_STATE,
      logs: prev.logs,
      showLogs: prev.showLogs
    }));
  }, []);

  // ── Tool Selection ──────────────────────────────────────────────

  const handleToolSelect = React.useCallback((_: unknown, option?: IDropdownOption): void => {
    if (option) {
      setState(prev => ({
        ...prev,
        selectedToolName: option.key as string,
        parameterValues: {},
        lastRequest: undefined,
        lastResponse: undefined,
        lastDurationMs: undefined
      }));
    }
  }, []);

  const handleParamChange = React.useCallback((paramName: string, value: string): void => {
    setState(prev => ({
      ...prev,
      parameterValues: { ...prev.parameterValues, [paramName]: value }
    }));
  }, []);

  // ── Preset helpers ───────────────────────────────────────────────

  const customExamplesForTool: IToolExample[] = state.selectedToolName
    ? (state.customPresets[state.selectedToolName] || [])
    : [];

  const handleExampleSelect = React.useCallback((_: unknown, option?: IDropdownOption): void => {
    if (!option || !state.selectedToolName) return;
    const key = option.key as string;

    // Find the example: built-in or custom
    let example: IToolExample | undefined;
    if (key.indexOf('builtin-') === 0) {
      const idx = parseInt(key.substring(8), 10);
      example = selectedExamples[idx];
    } else if (key.indexOf('custom-') === 0) {
      const idx = parseInt(key.substring(7), 10);
      example = customExamplesForTool[idx];
    }
    if (!example) return;

    const newParams: Record<string, string> = {};
    Object.keys(example.params).forEach(k => {
      const val = example!.params[k];
      newParams[k] = Array.isArray(val) ? val.join(', ') : String(val);
    });

    setState(prev => ({ ...prev, parameterValues: newParams }));
  }, [state.selectedToolName, selectedExamples, customExamplesForTool]);

  const handleSavePreset = React.useCallback((): void => {
    if (!state.selectedToolName || !state.presetSaveName.trim()) return;

    const preset: IToolExample = {
      label: state.presetSaveName.trim(),
      params: { ...state.parameterValues }
    };

    const updated = { ...state.customPresets };
    const toolPresets = updated[state.selectedToolName] || [];
    updated[state.selectedToolName] = [...toolPresets, preset];

    saveCustomPresets(updated);
    setState(prev => ({ ...prev, customPresets: updated, presetSaveName: '', showPresetSave: false }));
  }, [state.selectedToolName, state.presetSaveName, state.parameterValues, state.customPresets]);

  const handleDeletePreset = React.useCallback((toolName: string, idx: number): void => {
    const updated = { ...state.customPresets };
    const toolPresets = [...(updated[toolName] || [])];
    toolPresets.splice(idx, 1);
    if (toolPresets.length === 0) {
      delete updated[toolName];
    } else {
      updated[toolName] = toolPresets;
    }
    saveCustomPresets(updated);
    setState(prev => ({ ...prev, customPresets: updated }));
  }, [state.customPresets]);

  // ── Execute Tool ────────────────────────────────────────────────

  const handleExecute = React.useCallback(async (): Promise<void> => {
    if (!clientRef.current || !state.selectedToolName) return;

    const args: Record<string, unknown> = {};
    if (selectedSchema?.properties) {
      Object.keys(state.parameterValues).forEach(key => {
        const value = state.parameterValues[key];
        if (!value || value.trim() === '') return;

        const propDef = selectedSchema.properties[key];
        if (propDef && propDef.type === 'array') {
          // Try parsing as JSON first (for array of objects like [{Key, Value}])
          const trimmed = value.trim();
          if (trimmed.indexOf('[') === 0) {
            try {
              args[key] = JSON.parse(trimmed);
            } catch {
              args[key] = value.split(',').map(v => v.trim()).filter(v => v !== '');
            }
          } else if (trimmed.indexOf('{') === 0) {
            // Single object — wrap in array
            try {
              args[key] = [JSON.parse(trimmed)];
            } catch {
              args[key] = value.split(',').map(v => v.trim()).filter(v => v !== '');
            }
          } else {
            args[key] = value.split(',').map(v => v.trim()).filter(v => v !== '');
          }
        } else if (propDef && propDef.type === 'object') {
          // Try parsing as JSON object
          try {
            args[key] = JSON.parse(value.trim());
          } catch {
            args[key] = value;
          }
        } else {
          args[key] = value;
        }
      });
    }

    const request = { tool: state.selectedToolName, args };
    setState(prev => ({ ...prev, isExecuting: true, lastRequest: request, lastResponse: undefined, lastDurationMs: undefined }));

    const startTime = Date.now();
    try {
      const result = await clientRef.current.callTool(state.selectedToolName, args);
      setState(prev => ({
        ...prev,
        isExecuting: false,
        lastResponse: result,
        lastDurationMs: Date.now() - startTime
      }));
    } catch (err) {
      const errMsg = (err as Error).message || '';
      const isSessionExpired = errMsg.indexOf('HTTP 500') !== -1 || errMsg.indexOf('HTTP 401') !== -1 || errMsg.indexOf('Failed to fetch') !== -1;
      setState(prev => ({
        ...prev,
        isExecuting: false,
        connectionState: isSessionExpired ? 'error' : prev.connectionState,
        connectionError: isSessionExpired ? 'Session expired or server unavailable. Please disconnect and reconnect.' : prev.connectionError,
        lastResponse: { error: errMsg },
        lastDurationMs: Date.now() - startTime
      }));
    }
  }, [state.selectedToolName, state.parameterValues, selectedSchema]);

  // ── Execute Prompt ──────────────────────────────────────────────

  const handleGetPrompt = React.useCallback(async (promptName: string): Promise<void> => {
    if (!clientRef.current) return;
    setState(prev => ({ ...prev, isExecuting: true, lastRequest: { prompt: promptName }, lastResponse: undefined, lastDurationMs: undefined }));

    const startTime = Date.now();
    try {
      const result = await clientRef.current.getPrompt(promptName);
      setState(prev => ({ ...prev, isExecuting: false, lastResponse: result, lastDurationMs: Date.now() - startTime }));
    } catch (err) {
      setState(prev => ({ ...prev, isExecuting: false, lastResponse: { error: (err as Error).message }, lastDurationMs: Date.now() - startTime }));
    }
  }, []);

  // ── Read Resource ───────────────────────────────────────────────

  const handleReadResource = React.useCallback(async (uri: string): Promise<void> => {
    if (!clientRef.current) return;
    setState(prev => ({ ...prev, isExecuting: true, lastRequest: { resource: uri }, lastResponse: undefined, lastDurationMs: undefined }));

    const startTime = Date.now();
    try {
      const result = await clientRef.current.readResource(uri);
      setState(prev => ({ ...prev, isExecuting: false, lastResponse: result, lastDurationMs: Date.now() - startTime }));
    } catch (err) {
      setState(prev => ({ ...prev, isExecuting: false, lastResponse: { error: (err as Error).message }, lastDurationMs: Date.now() - startTime }));
    }
  }, []);

  // ── Styles ──────────────────────────────────────────────────────

  const toolOptions: IDropdownOption[] = [...toolList]
    .sort((a, b) => a.name.localeCompare(b.name))
    .map(t => ({
      key: t.name,
      text: t.name,
      title: t.description
    }));

  const exampleOptions: IDropdownOption[] = [];
  if (selectedExamples.length > 0) {
    exampleOptions.push({ key: 'header-builtin', text: 'Built-in examples (educational)', itemType: DropdownMenuItemType.Header });
    selectedExamples.forEach((ex, i) => {
      exampleOptions.push({ key: `builtin-${i}`, text: ex.label });
    });
  }
  if (customExamplesForTool.length > 0) {
    if (selectedExamples.length > 0) {
      exampleOptions.push({ key: 'divider', text: '-', itemType: DropdownMenuItemType.Divider });
    }
    exampleOptions.push({ key: 'header-custom', text: 'Your presets (browser localStorage)', itemType: DropdownMenuItemType.Header });
    customExamplesForTool.forEach((ex, i) => {
      exampleOptions.push({ key: `custom-${i}`, text: ex.label });
    });
  }

  const panelStyle: React.CSSProperties = {
    border: `1px solid ${theme.palette.neutralLight}`,
    borderRadius: 4,
    padding: 12,
    backgroundColor: theme.palette.white
  };

  const codeBlockStyle: React.CSSProperties = {
    fontFamily: 'Consolas, Monaco, "Courier New", monospace',
    fontSize: 12,
    backgroundColor: theme.palette.neutralLighterAlt,
    padding: 12,
    borderRadius: 4,
    overflow: 'auto',
    maxHeight: 400,
    whiteSpace: 'pre-wrap',
    wordBreak: 'break-word'
  };

  const capBadgeStyle = (active: boolean): React.CSSProperties => ({
    display: 'inline-block',
    padding: '2px 8px',
    borderRadius: 12,
    fontSize: 11,
    fontWeight: 600,
    marginRight: 6,
    backgroundColor: active ? theme.palette.themeLighter : theme.palette.neutralLighter,
    color: active ? theme.palette.themeDarkAlt : theme.palette.neutralTertiary,
    border: `1px solid ${active ? theme.palette.themeLight : theme.palette.neutralQuaternary}`
  });

  const isConnected = state.connectionState === 'connected';
  const capKeys = state.capabilities ? Object.keys(state.capabilities) : [];

  // ── Render ──────────────────────────────────────────────────────

  return (
    <div style={{ padding: 16, fontFamily: theme.fonts.medium.fontFamily }}>
      {/* ── Header ────────────────────────────────────────────── */}
      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 12 }} styles={{ root: { marginBottom: 16 } }}>
        <Icon iconName="ContactCard" styles={{ root: { fontSize: 24, color: theme.palette.themePrimary } }} />
        <Text variant="xLarge" styles={{ root: { fontWeight: 600 } }}>MCP365 Explorer: User Profile</Text>
        <span style={{
          fontFamily: 'Consolas, Monaco, "Courier New", monospace',
          fontSize: 12,
          fontWeight: 600,
          color: theme.palette.themeDarkAlt,
          backgroundColor: theme.palette.themeLighter,
          padding: '2px 8px',
          borderRadius: 4,
          border: `1px solid ${theme.palette.themeLight}`
        }}>
          mcp_MeServer
        </span>
      </Stack>

      {/* ── Connection Bar ────────────────────────────────────── */}
      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }} styles={{ root: { marginBottom: 16 } }}>
        {state.connectionState === 'disconnected' || state.connectionState === 'error' ? (
          <PrimaryButton text="Connect" iconProps={{ iconName: 'PlugConnected' }} onClick={handleConnect} disabled={!props.environmentId} />
        ) : state.connectionState === 'connecting' ? (
          <Spinner size={SpinnerSize.small} label="Connecting..." />
        ) : (
          <DefaultButton text="Disconnect" iconProps={{ iconName: 'PlugDisconnected' }} onClick={handleDisconnect} />
        )}
        {isConnected && (
          <MessageBar messageBarType={MessageBarType.success} isMultiline={false}>
            Connected &middot; {state.discoveredTools.length} tools, {state.discoveredPrompts.length} prompts, {state.discoveredResources.length} resources
          </MessageBar>
        )}
        {!props.environmentId && (
          <MessageBar messageBarType={MessageBarType.warning} isMultiline={false}>
            Configure the Environment ID in the web part property pane to connect.
          </MessageBar>
        )}
      </Stack>

      {state.connectionError && (
        <MessageBar messageBarType={MessageBarType.error} styles={{ root: { marginBottom: 12 } }}>
          {state.connectionError}
        </MessageBar>
      )}

      {/* ── Server Info Panel ─────────────────────────────────── */}
      {isConnected && state.serverInfo && (
        <div style={{ ...panelStyle, marginBottom: 16 }}>
          <Stack horizontal tokens={{ childrenGap: 24 }} verticalAlign="center">
            <Stack tokens={{ childrenGap: 2 }}>
              <Text variant="smallPlus" styles={{ root: { fontWeight: 600 } }}>Server</Text>
              <Text variant="small">{state.serverInfo.name} v{state.serverInfo.version}</Text>
            </Stack>
            <Stack tokens={{ childrenGap: 2 }}>
              <Text variant="smallPlus" styles={{ root: { fontWeight: 600 } }}>Protocol</Text>
              <Text variant="small">{state.protocolVersion}</Text>
            </Stack>
            <Stack tokens={{ childrenGap: 2 }}>
              <Text variant="smallPlus" styles={{ root: { fontWeight: 600 } }}>Session</Text>
              <TooltipHost content={state.sessionId || 'No session ID'}>
                <Text variant="small" styles={{ root: { fontFamily: 'Consolas, monospace', cursor: 'default' } }}>
                  {state.sessionId ? (state.sessionId.length > 20 ? state.sessionId.substring(0, 20) + '...' : state.sessionId) : 'none'}
                </Text>
              </TooltipHost>
            </Stack>
            <Stack tokens={{ childrenGap: 2 }}>
              <Text variant="smallPlus" styles={{ root: { fontWeight: 600 } }}>Capabilities</Text>
              <div>
                {['tools', 'prompts', 'resources', 'logging'].map(cap => (
                  <span key={cap} style={capBadgeStyle(capKeys.indexOf(cap) !== -1)}>
                    {cap}
                  </span>
                ))}
              </div>
            </Stack>
          </Stack>
        </div>
      )}

      {/* ── Main Pivot: Showcase / Tools / Prompts / Resources ── */}
      <Pivot styles={{ root: { marginBottom: 16 } }} selectedKey={state.activeTab} onLinkClick={(item) => { if (item?.props.itemKey) setState(prev => ({ ...prev, activeTab: item.props.itemKey as string })); }}>

        {/* ── Showcase Tab (only when connected) ─────────────── */}
        {isConnected && clientRef.current && (
          <PivotItem headerText="Showcase" itemIcon="Rocket" itemKey="showcase">
            <UserProfileShowcase client={clientRef.current} theme={theme} onLog={addLog} />
          </PivotItem>
        )}

        {/* ── Tools Tab ──────────────────────────────────────── */}
        <PivotItem headerText={`Tools (${toolList.length})`} itemIcon="Settings" itemKey="tools">
          {/* Source indicator — only when connected */}
          {isConnected && toolList.length > 0 && (
            <div style={{ marginTop: 8, marginBottom: 8 }}>
              <span style={{
                display: 'inline-block',
                padding: '2px 10px',
                borderRadius: 12,
                fontSize: 11,
                fontWeight: 600,
                backgroundColor: '#dff6dd',
                color: '#107c10',
                border: '1px solid #a7e3a5'
              }}>
                Live from server (tools/list)
              </span>
            </div>
          )}
          {!isConnected && (
            <Text variant="small" styles={{ root: { marginTop: 8, marginBottom: 8, color: theme.palette.neutralTertiary } }}>
              Connect to discover and test tools.
            </Text>
          )}

          <Stack horizontal tokens={{ childrenGap: 16 }}>
            {/* Left: Tool Selection + Description + Schema */}
            <Stack style={{ minWidth: 280, maxWidth: 340, ...panelStyle }}>
              <Label>Tool</Label>
              <Dropdown
                options={toolOptions}
                selectedKey={state.selectedToolName}
                onChange={handleToolSelect}
                placeholder={isConnected ? 'Select a tool' : 'Connect first'}
                disabled={!isConnected}
              />

              {/* Tool description from source */}
              {selectedTool && (
                <div style={{
                  marginTop: 10,
                  padding: 8,
                  backgroundColor: '#f6faf6',
                  borderRadius: 4,
                  borderLeft: `3px solid ${'#107c10'}`
                }}>
                  <Text variant="small" styles={{ root: { fontWeight: 600, display: 'block', marginBottom: 4 } }}>
                    Description (from server)
                  </Text>
                  <Text variant="small" styles={{ root: { color: theme.palette.neutralPrimary, lineHeight: '1.4' } }}>
                    {selectedTool.description}
                  </Text>
                </div>
              )}

              {/* Examples & Presets */}
              {(exampleOptions.length > 0 || state.selectedToolName) && (
                <>
                  <Label styles={{ root: { marginTop: 12 } }}>Examples &amp; Presets</Label>
                  {exampleOptions.length > 0 && (
                    <Dropdown
                      options={exampleOptions}
                      placeholder="Load example or preset"
                      onChange={handleExampleSelect}
                    />
                  )}
                  {customExamplesForTool.length > 0 && (
                    <Stack tokens={{ childrenGap: 2 }} styles={{ root: { marginTop: 4 } }}>
                      {customExamplesForTool.map((preset, idx) => (
                        <Stack key={idx} horizontal verticalAlign="center" tokens={{ childrenGap: 4 }}>
                          <Text variant="small" styles={{ root: { color: theme.palette.neutralSecondary, flex: 1 } }}>
                            {preset.label}
                          </Text>
                          <DefaultButton
                            iconProps={{ iconName: 'Delete' }}
                            title="Delete this preset"
                            styles={{ root: { minWidth: 0, width: 24, height: 20, padding: 0, border: 'none' }, icon: { fontSize: 10 } }}
                            onClick={() => { if (state.selectedToolName) handleDeletePreset(state.selectedToolName, idx); }}
                          />
                        </Stack>
                      ))}
                    </Stack>
                  )}
                  {/* Save preset UI */}
                  {state.showPresetSave ? (
                    <Stack horizontal tokens={{ childrenGap: 4 }} styles={{ root: { marginTop: 6 } }}>
                      <TextField
                        placeholder="Preset name"
                        value={state.presetSaveName}
                        onChange={(_, val) => setState(prev => ({ ...prev, presetSaveName: val || '' }))}
                        styles={{ root: { flex: 1 } }}
                        onKeyDown={(e) => { if (e.key === 'Enter') handleSavePreset(); }}
                      />
                      <PrimaryButton text="Save" onClick={handleSavePreset} disabled={!state.presetSaveName.trim()} styles={{ root: { minWidth: 0 } }} />
                      <DefaultButton text="Cancel" onClick={() => setState(prev => ({ ...prev, showPresetSave: false, presetSaveName: '' }))} styles={{ root: { minWidth: 0 } }} />
                    </Stack>
                  ) : (
                    <DefaultButton
                      text="Save current as preset"
                      iconProps={{ iconName: 'Save' }}
                      onClick={() => setState(prev => ({ ...prev, showPresetSave: true }))}
                      disabled={!state.selectedToolName || Object.keys(state.parameterValues).length === 0}
                      styles={{ root: { marginTop: 6 } }}
                    />
                  )}
                  {/* Storage info */}
                  <Text variant="small" styles={{ root: { marginTop: 6, color: theme.palette.neutralTertiary, fontStyle: 'italic' } }}>
                    {customExamplesForTool.length > 0
                      ? 'Your presets are saved in browser localStorage'
                      : 'Save your own parameter presets to browser localStorage'}
                  </Text>
                </>
              )}

              {/* Input Schema — raw JSON from source */}
              {selectedSchema && (
                <>
                  <Label styles={{ root: { marginTop: 12 } }}>
                    Input Schema (from server)
                  </Label>
                  <div style={{ ...codeBlockStyle, maxHeight: 250, fontSize: 11 }}>
                    {JSON.stringify(selectedSchema, null, 2)}
                  </div>
                </>
              )}
            </Stack>

            {/* Right: Parameters + Execute */}
            <Stack grow style={{ ...panelStyle, flex: 1 }}>
              <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                <Label>Parameters</Label>
                {isConnected && (
                  <Text variant="small" styles={{ root: { color: theme.palette.neutralTertiary, fontStyle: 'italic' } }}>
                    Field descriptions come from the live server schema
                  </Text>
                )}
              </Stack>
              {selectedSchema?.properties && (
                <Stack tokens={{ childrenGap: 8 }}>
                  {Object.keys(selectedSchema.properties).map(paramName => {
                    const propDef = selectedSchema.properties[paramName];
                    const isRequired = (selectedSchema.required || []).indexOf(paramName) !== -1;
                    return (
                      <TextField
                        key={paramName}
                        label={`${paramName}${isRequired ? ' *' : ''}`}
                        description={propDef.description}
                        placeholder={propDef.type === 'array' ? 'JSON array or comma-separated values' : propDef.type === 'object' ? 'JSON object' : propDef.type}
                        multiline={propDef.type === 'array' || propDef.type === 'object'}
                        rows={propDef.type === 'array' || propDef.type === 'object' ? 3 : undefined}
                        value={state.parameterValues[paramName] || ''}
                        onChange={(_, val) => handleParamChange(paramName, val || '')}
                        required={isRequired}
                      />
                    );
                  })}
                </Stack>
              )}
              <Stack horizontal tokens={{ childrenGap: 8 }} styles={{ root: { marginTop: 12 } }}>
                <PrimaryButton
                  text={state.isExecuting ? 'Executing...' : 'Execute'}
                  iconProps={{ iconName: 'Play' }}
                  onClick={handleExecute}
                  disabled={state.isExecuting || !isConnected || !state.selectedToolName}
                />
                {state.lastDurationMs !== undefined && (
                  <Text variant="small" styles={{ root: { alignSelf: 'center', color: theme.palette.neutralSecondary } }}>
                    {state.lastDurationMs}ms
                  </Text>
                )}
              </Stack>
            </Stack>
          </Stack>
        </PivotItem>

        {/* ── Prompts Tab ────────────────────────────────────── */}
        <PivotItem headerText={`Prompts (${state.discoveredPrompts.length})`} itemIcon="TextDocument">
          <div style={{ ...panelStyle, marginTop: 12 }}>
            {!isConnected ? (
              <Text variant="small" styles={{ root: { color: theme.palette.neutralTertiary } }}>Connect to discover prompts.</Text>
            ) : state.discoveredPrompts.length === 0 ? (
              <Text variant="small" styles={{ root: { color: theme.palette.neutralTertiary } }}>
                No prompts discovered. The server advertises the prompts capability but returned no prompt templates.
              </Text>
            ) : (
              <Stack tokens={{ childrenGap: 8 }}>
                {state.discoveredPrompts.map(prompt => (
                  <div key={prompt.name} style={{ ...panelStyle, padding: 8 }}>
                    <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                      <Icon iconName="TextDocument" styles={{ root: { color: theme.palette.themePrimary } }} />
                      <Text styles={{ root: { fontWeight: 600 } }}>{prompt.name}</Text>
                      <DefaultButton
                        text="Get"
                        styles={{ root: { minWidth: 0, padding: '0 8px', height: 24 } }}
                        onClick={() => { void handleGetPrompt(prompt.name); }}
                        disabled={state.isExecuting}
                      />
                    </Stack>
                    {prompt.description && (
                      <Text variant="small" styles={{ root: { color: theme.palette.neutralSecondary, marginTop: 4 } }}>
                        {prompt.description}
                      </Text>
                    )}
                    {prompt.arguments && prompt.arguments.length > 0 && (
                      <div style={{ marginTop: 4 }}>
                        <Text variant="small" styles={{ root: { fontWeight: 600 } }}>Arguments: </Text>
                        <Text variant="small">
                          {prompt.arguments.map(a => `${a.name}${a.required ? '*' : ''}`).join(', ')}
                        </Text>
                      </div>
                    )}
                  </div>
                ))}
              </Stack>
            )}
          </div>
        </PivotItem>

        {/* ── Resources Tab ──────────────────────────────────── */}
        <PivotItem headerText={`Resources (${state.discoveredResources.length})`} itemIcon="Database">
          <div style={{ ...panelStyle, marginTop: 12 }}>
            {!isConnected ? (
              <Text variant="small" styles={{ root: { color: theme.palette.neutralTertiary } }}>Connect to discover resources.</Text>
            ) : state.discoveredResources.length === 0 ? (
              <Text variant="small" styles={{ root: { color: theme.palette.neutralTertiary } }}>
                No resources discovered. The server advertises the resources capability but returned no resource endpoints.
              </Text>
            ) : (
              <Stack tokens={{ childrenGap: 8 }}>
                {state.discoveredResources.map(resource => (
                  <div key={resource.uri} style={{ ...panelStyle, padding: 8 }}>
                    <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                      <Icon iconName="Database" styles={{ root: { color: theme.palette.themePrimary } }} />
                      <Text styles={{ root: { fontWeight: 600 } }}>{resource.name}</Text>
                      {resource.mimeType && (
                        <Text variant="small" styles={{ root: { color: theme.palette.neutralTertiary } }}>({resource.mimeType})</Text>
                      )}
                      <DefaultButton
                        text="Read"
                        styles={{ root: { minWidth: 0, padding: '0 8px', height: 24 } }}
                        onClick={() => { void handleReadResource(resource.uri); }}
                        disabled={state.isExecuting}
                      />
                    </Stack>
                    {resource.description && (
                      <Text variant="small" styles={{ root: { color: theme.palette.neutralSecondary, marginTop: 4 } }}>
                        {resource.description}
                      </Text>
                    )}
                    <Text variant="small" styles={{ root: { fontFamily: 'Consolas, monospace', color: theme.palette.neutralTertiary, marginTop: 2 } }}>
                      {resource.uri}
                    </Text>
                  </div>
                ))}
              </Stack>
            )}
          </div>
        </PivotItem>

        {/* ── Raw Capabilities Tab ────────────────────────────── */}
        <PivotItem headerText="Raw" itemIcon="Code">
          <div style={{ ...codeBlockStyle, marginTop: 12 }}>
            {isConnected ? JSON.stringify({
              protocolVersion: state.protocolVersion,
              serverInfo: state.serverInfo,
              capabilities: state.capabilities,
              sessionId: state.sessionId,
              toolCount: state.discoveredTools.length,
              promptCount: state.discoveredPrompts.length,
              resourceCount: state.discoveredResources.length
            }, null, 2) : 'Connect to see raw server data.'}
          </div>
        </PivotItem>
      </Pivot>

      {/* ── Results ───────────────────────────────────────────── */}
      {(state.lastRequest || state.lastResponse) && (
        <div style={{ ...panelStyle, marginBottom: 16 }}>
          <Pivot>
            <PivotItem headerText="Formatted" itemIcon="TextDocument">
              <div style={{ ...codeBlockStyle, padding: 0, backgroundColor: 'transparent' }}>
                {state.lastResponse
                  ? <FormattedResponse response={state.lastResponse} theme={theme} />
                  : state.isExecuting
                    ? <div style={{ padding: 12, color: theme.palette.neutralTertiary }}>Executing...</div>
                    : <div style={{ padding: 12, color: theme.palette.neutralTertiary }}>No response yet</div>}
              </div>
            </PivotItem>
            <PivotItem headerText="Raw Response" itemIcon="Code">
              {state.lastResponse && (
                <DefaultButton
                  text="Copy"
                  iconProps={{ iconName: 'Copy' }}
                  onClick={() => { void navigator.clipboard.writeText(JSON.stringify(state.lastResponse, null, 2)); }}
                  styles={{ root: { minWidth: 0, padding: '0 8px', marginBottom: 4, marginTop: 4 } }}
                />
              )}
              <div style={codeBlockStyle}>
                {state.lastResponse
                  ? JSON.stringify(state.lastResponse, null, 2)
                  : state.isExecuting ? 'Executing...' : 'No response yet'}
              </div>
            </PivotItem>
            <PivotItem headerText="Request" itemIcon="Send">
              {state.lastRequest && (
                <DefaultButton
                  text="Copy"
                  iconProps={{ iconName: 'Copy' }}
                  onClick={() => { void navigator.clipboard.writeText(JSON.stringify(state.lastRequest, null, 2)); }}
                  styles={{ root: { minWidth: 0, padding: '0 8px', marginBottom: 4, marginTop: 4 } }}
                />
              )}
              <div style={codeBlockStyle}>
                {state.lastRequest ? JSON.stringify(state.lastRequest, null, 2) : ''}
              </div>
            </PivotItem>
          </Pivot>
        </div>
      )}

      {/* ── Log Viewer ────────────────────────────────────────── */}
      <div style={panelStyle}>
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
          <DefaultButton
            text={state.showLogs ? 'Hide Logs' : 'Show Logs'}
            iconProps={{ iconName: state.showLogs ? 'ChevronDown' : 'ChevronRight' }}
            onClick={() => setState(prev => ({ ...prev, showLogs: !prev.showLogs }))}
            styles={{ root: { minWidth: 0, padding: '0 8px' } }}
          />
          <Text variant="small" styles={{ root: { color: theme.palette.neutralSecondary } }}>
            {state.logs.length} entries
          </Text>
          {state.logs.length > 0 && (
            <DefaultButton
              text="Clear"
              iconProps={{ iconName: 'Delete' }}
              onClick={() => setState(prev => ({ ...prev, logs: [], logExpandedIndex: undefined }))}
              styles={{ root: { minWidth: 0, padding: '0 8px' } }}
            />
          )}
        </Stack>
        {state.showLogs && (
          <LogViewer
            logs={state.logs}
            filter={state.logFilter}
            expandedIndex={state.logExpandedIndex}
            onFilterChange={(val) => setState(prev => ({ ...prev, logFilter: val }))}
            onToggleExpand={(idx) => setState(prev => ({ ...prev, logExpandedIndex: prev.logExpandedIndex === idx ? undefined : idx }))}
            theme={theme}
          />
        )}
      </div>
    </div>
  );
};

