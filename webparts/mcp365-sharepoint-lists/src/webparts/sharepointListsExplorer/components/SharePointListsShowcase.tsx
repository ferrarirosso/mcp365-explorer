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
  MessageBar,
  MessageBarType,
  getTheme
} from '@fluentui/react';
import { FilePicker, IFilePickerResult } from '@pnp/spfx-controls-react/lib/FilePicker';
import type { BaseComponentContext } from '@microsoft/sp-component-base';
import type { McpBrowserClient, IMcpCallResult, McpLogHandler } from '../services/McpBrowserClient';

// ─── Types ──────────────────────────────────────────────────────────

interface IShowcaseProps {
  client: McpBrowserClient;
  theme: ReturnType<typeof getTheme>;
  onLog: McpLogHandler;
  userEmail: string;
  spfxContext: BaseComponentContext;
}

type StepStatus = 'pending' | 'running' | 'done' | 'error';

interface IShowcaseState {
  docLibId: string | undefined;
  docLibName: string | undefined;
  docLibWebUrl: string | undefined;
  folderId: string | undefined;
  folderName: string;
  folderNameInput: string;
  folderWebUrl: string | undefined;
  uploadedFileName: string | undefined;
  uploadStatus: string | undefined;
  operationToken: string | undefined;
  operationResult: string | undefined;
  selectedFileUrl: string | undefined;

  stepStatus: {
    getDocLib: StepStatus;
    createFolder: StepStatus;
    uploadFile: StepStatus;
    checkStatus: StepStatus;
  };

  loading: string | undefined;
  error: string | undefined;
}

const INITIAL: IShowcaseState = {
  docLibId: undefined,
  docLibName: undefined,
  docLibWebUrl: undefined,
  folderId: undefined,
  folderName: '',
  folderNameInput: 'MCP Explorer Demo',
  folderWebUrl: undefined,
  uploadedFileName: undefined,
  uploadStatus: undefined,
  operationToken: undefined,
  operationResult: undefined,
  selectedFileUrl: undefined,
  stepStatus: {
    getDocLib: 'pending',
    createFolder: 'pending',
    uploadFile: 'pending',
    checkStatus: 'pending'
  },
  loading: undefined,
  error: undefined
};

// ─── Helpers ────────────────────────────────────────────────────────

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

function checkMcpError(result: IMcpCallResult): void {
  if (result.isError && result.content && result.content.length > 0) {
    throw new Error(result.content[0].text || 'MCP tool returned an error');
  }
  // Also check for inline error objects
  const parsed = extractJsonFromContent(result);
  if (parsed && typeof parsed === 'object' && (parsed as Record<string, unknown>).Error === 'Tool Request Failed') {
    const errObj = parsed as Record<string, unknown>;
    throw new Error(`${errObj.StatusCode} ${errObj.StatusDescription}`);
  }
}

// ─── Component ──────────────────────────────────────────────────────

export const SharePointListsShowcase: React.FC<IShowcaseProps> = ({ client, theme, spfxContext }) => {
  const [state, setState] = React.useState<IShowcaseState>(INITIAL);

  const setStep = (step: keyof IShowcaseState['stepStatus'], status: StepStatus): void => {
    setState(prev => ({ ...prev, stepStatus: { ...prev.stepStatus, [step]: status } }));
  };

  // ── Step order for active indicator ───────────────────────────

  const s = state.stepStatus;
  const stepOrder: Array<keyof typeof s> = ['getDocLib', 'createFolder', 'uploadFile', 'checkStatus'];
  const activeStepIndex = stepOrder.findIndex(k => s[k] !== 'done');
  const activeStepKey = activeStepIndex >= 0 ? stepOrder[activeStepIndex] : undefined;

  const stepIcon = (status: StepStatus, stepKey: keyof typeof s): string => {
    if (status === 'done') return 'SkypeCircleCheck';
    if (status === 'error') return 'ErrorBadge';
    if (status === 'running') return 'Sync';
    if (stepKey === activeStepKey) return 'LocationDot';
    return 'CircleRing';
  };

  const stepColor = (status: StepStatus, stepKey: keyof typeof s): string => {
    if (status === 'done') return '#107c10';
    if (status === 'error') return theme.palette.red;
    if (status === 'running') return theme.palette.themePrimary;
    if (stepKey === activeStepKey) return theme.palette.themePrimary;
    return theme.palette.neutralTertiary;
  };

  const stepStyle = (status: StepStatus): React.CSSProperties => ({
    border: `1px solid ${status === 'done' ? '#a7e3a5' : status === 'error' ? theme.palette.red : theme.palette.neutralLight}`,
    borderRadius: 8,
    padding: 16,
    backgroundColor: status === 'done' ? '#f6faf6' : theme.palette.white,
    marginBottom: 12
  });

  const isLoading = !!state.loading;

  // ── Step 1: Get default document library ──────────────────────

  const handleGetDocLib = React.useCallback(async (): Promise<void> => {
    setState(prev => ({ ...prev, loading: 'getDefaultDocumentLibraryInSite', error: undefined }));
    setStep('getDocLib', 'running');

    try {
      // Use the current site where the webpart is hosted
      const pageCtx = spfxContext.pageContext;
      const hostname = new URL(pageCtx.web.absoluteUrl).hostname;
      const currentSiteId = `${hostname},${pageCtx.site.id},${pageCtx.web.id}`;
      const result = await client.callTool('getDefaultDocumentLibraryInSite', { siteId: currentSiteId });
      checkMcpError(result);
      const parsed = extractJsonFromContent(result) as Record<string, unknown> | undefined;
      if (!parsed?.id) {
        setStep('getDocLib', 'error');
        setState(prev => ({ ...prev, loading: undefined, error: 'Could not extract document library ID.' }));
        return;
      }
      setState(prev => ({
        ...prev,
        loading: undefined,
        docLibId: String(parsed.id),
        docLibName: String(parsed.name || 'Documents'),
        docLibWebUrl: String(parsed.webUrl || '')
      }));
      setStep('getDocLib', 'done');
    } catch (err) {
      setStep('getDocLib', 'error');
      setState(prev => ({ ...prev, loading: undefined, error: (err as Error).message }));
    }
  }, [client]);

  // ── Step 2: Create folder ─────────────────────────────────────

  const handleCreateFolder = React.useCallback(async (): Promise<void> => {
    if (!state.docLibId || !state.folderNameInput.trim()) return;
    setState(prev => ({ ...prev, loading: 'createFolder', error: undefined }));
    setStep('createFolder', 'running');

    try {
      const result = await client.callTool('createFolder', {
        folderName: state.folderNameInput.trim(),
        documentLibraryId: state.docLibId
      });
      checkMcpError(result);
      const parsed = extractJsonFromContent(result) as Record<string, unknown> | undefined;
      setState(prev => ({
        ...prev,
        loading: undefined,
        folderId: parsed?.id ? String(parsed.id) : undefined,
        folderName: state.folderNameInput.trim(),
        folderWebUrl: parsed?.webUrl ? String(parsed.webUrl) : undefined
      }));
      setStep('createFolder', 'done');
    } catch (err) {
      setStep('createFolder', 'error');
      setState(prev => ({ ...prev, loading: undefined, error: (err as Error).message }));
    }
  }, [client, state.docLibId, state.folderNameInput]);

  // ── Step 3: Upload file ───────────────────────────────────────

  const handleFileSelected = React.useCallback((results: IFilePickerResult[]): void => {
    if (results.length > 0 && results[0].fileAbsoluteUrl) {
      setState(prev => ({ ...prev, selectedFileUrl: results[0].fileAbsoluteUrl }));
    }
  }, []);

  const handleUploadFile = React.useCallback(async (): Promise<void> => {
    if (!state.docLibId || !state.selectedFileUrl) return;
    setState(prev => ({ ...prev, loading: 'uploadFileFromUrl', error: undefined }));
    setStep('uploadFile', 'running');

    try {
      const params: Record<string, unknown> = {
        sourceUrl: state.selectedFileUrl,
        destinationDocumentLibraryId: state.docLibId
      };
      if (state.folderId) {
        params.destinationFolderId = state.folderId;
      }
      const result = await client.callTool('uploadFileFromUrl', params);
      checkMcpError(result);
      const parsed = extractJsonFromContent(result) as Record<string, unknown> | undefined;
      const status = parsed?.Status ? String(parsed.Status) : 'Completed';
      const message = parsed?.Message ? String(parsed.Message) : '';
      const opToken = parsed?.OperationToken ? String(parsed.OperationToken) : undefined;
      setState(prev => ({
        ...prev,
        loading: undefined,
        uploadedFileName: parsed?.name ? String(parsed.name) : (status === 'Accepted' ? 'Copy operation initiated' : 'File uploaded'),
        uploadStatus: message || status,
        operationToken: opToken
      }));
      setStep('uploadFile', 'done');
    } catch (err) {
      setStep('uploadFile', 'error');
      setState(prev => ({ ...prev, loading: undefined, error: (err as Error).message }));
    }
  }, [client, state.docLibId, state.folderId, state.selectedFileUrl]);

  // ── Render ────────────────────────────────────────────────────

  return (
    <div style={{ marginTop: 12 }}>
      <Text variant="small" styles={{ root: { color: theme.palette.neutralTertiary, marginBottom: 12, display: 'block' } }}>
        Explore the Work IQ SharePoint server: get the default document library, create a folder, and upload a file.
      </Text>

      {state.error && (
        <MessageBar messageBarType={MessageBarType.error} styles={{ root: { marginBottom: 12 } }}>
          {state.error}
        </MessageBar>
      )}

      {/* ── Step 1: Get Document Library ─────────────────────── */}
      <div style={stepStyle(s.getDocLib)}>
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
          <Icon iconName={stepIcon(s.getDocLib, 'getDocLib')} styles={{ root: { color: stepColor(s.getDocLib, 'getDocLib'), fontSize: 16 } }} />
          <Text styles={{ root: { fontWeight: 600 } }}>Step 1: Get default document library</Text>
          <Text variant="small" styles={{ root: { color: theme.palette.neutralTertiary } }}>(getDefaultDocumentLibraryInSite)</Text>
        </Stack>
        <PrimaryButton
          text={state.loading === 'getDefaultDocumentLibraryInSite' ? 'Loading...' : 'Get Document Library'}
          iconProps={{ iconName: 'FabricFolder' }}
          onClick={handleGetDocLib}
          disabled={isLoading || s.getDocLib === 'done'}
          styles={{ root: { marginTop: 8 } }}
        />
        {state.docLibName && (
          <Stack tokens={{ childrenGap: 2 }} styles={{ root: { marginTop: 8 } }}>
            <Text variant="small" styles={{ root: { color: '#107c10', fontWeight: 600 } }}>
              {state.docLibName}
            </Text>
            <Text variant="small" styles={{ root: { fontFamily: 'Consolas, monospace', color: theme.palette.neutralSecondary, fontSize: 11 } }}>
              ID: {state.docLibId}
            </Text>
            {state.docLibWebUrl && (
              <Text variant="small" styles={{ root: { color: theme.palette.neutralTertiary } }}>
                {state.docLibWebUrl}
              </Text>
            )}
          </Stack>
        )}
      </div>

      {/* ── Step 2: Create Folder ────────────────────────────── */}
      <div style={stepStyle(s.createFolder)}>
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
          <Icon iconName={stepIcon(s.createFolder, 'createFolder')} styles={{ root: { color: stepColor(s.createFolder, 'createFolder'), fontSize: 16 } }} />
          <Text styles={{ root: { fontWeight: 600 } }}>Step 2: Create a folder</Text>
          <Text variant="small" styles={{ root: { color: theme.palette.neutralTertiary } }}>(createFolder)</Text>
        </Stack>
        <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="end" styles={{ root: { marginTop: 8 } }}>
          <TextField
            label="Folder name"
            value={state.folderNameInput}
            onChange={(_, val) => setState(prev => ({ ...prev, folderNameInput: val || '' }))}
            styles={{ root: { width: 300 } }}
            disabled={s.getDocLib !== 'done' || s.createFolder === 'done'}
          />
          <DefaultButton
            text={state.loading === 'createFolder' ? 'Creating...' : 'Create'}
            iconProps={{ iconName: 'NewFolder' }}
            onClick={handleCreateFolder}
            disabled={isLoading || s.getDocLib !== 'done' || s.createFolder === 'done' || !state.folderNameInput.trim()}
          />
        </Stack>
        {state.folderName && s.createFolder === 'done' && (
          <Stack tokens={{ childrenGap: 2 }} styles={{ root: { marginTop: 4 } }}>
            <Text variant="small" styles={{ root: { color: '#107c10', fontWeight: 600 } }}>
              Folder &ldquo;{state.folderName}&rdquo; created
            </Text>
            {state.folderWebUrl && (
              <a href={state.folderWebUrl} target="_blank" rel="noopener noreferrer" style={{ fontSize: 12, color: theme.palette.themePrimary }}>
                {decodeURIComponent(state.folderWebUrl)}
              </a>
            )}
          </Stack>
        )}
      </div>

      {/* ── Step 3: Upload File ──────────────────────────────── */}
      <div style={stepStyle(s.uploadFile)}>
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
          <Icon iconName={stepIcon(s.uploadFile, 'uploadFile')} styles={{ root: { color: stepColor(s.uploadFile, 'uploadFile'), fontSize: 16 } }} />
          <Text styles={{ root: { fontWeight: 600 } }}>Step 3: Upload a file to the folder</Text>
          <Text variant="small" styles={{ root: { color: theme.palette.neutralTertiary } }}>(uploadFileFromUrl)</Text>
        </Stack>
        <Stack tokens={{ childrenGap: 8 }} styles={{ root: { marginTop: 8 } }}>
          <FilePicker
            context={spfxContext as never}
            buttonLabel="Pick a file"
            buttonIcon="CloudUpload"
            onSave={handleFileSelected}
            onChange={handleFileSelected}
            hideLinkUploadTab={true}
            hideLocalUploadTab={true}
            hideLocalMultipleUploadTab={true}
            disabled={s.createFolder !== 'done' || s.uploadFile === 'done'}
          />
          {state.selectedFileUrl && (
            <Text variant="small" styles={{ root: { color: theme.palette.neutralSecondary, fontFamily: 'Consolas, monospace', fontSize: 11 } }}>
              Selected: {state.selectedFileUrl}
            </Text>
          )}
          <DefaultButton
            text={state.loading === 'uploadFileFromUrl' ? 'Uploading...' : 'Upload to folder'}
            iconProps={{ iconName: 'Upload' }}
            onClick={handleUploadFile}
            disabled={isLoading || !state.selectedFileUrl || s.uploadFile === 'done'}
          />
        </Stack>
        {s.uploadFile === 'done' && (
          <Stack tokens={{ childrenGap: 2 }} styles={{ root: { marginTop: 4 } }}>
            <Text variant="small" styles={{ root: { color: '#107c10', fontWeight: 600 } }}>
              {state.uploadedFileName}
            </Text>
            {state.uploadStatus && (
              <Text variant="small" styles={{ root: { color: theme.palette.neutralSecondary } }}>
                {state.uploadStatus}
              </Text>
            )}
            {state.folderWebUrl && (
              <a href={state.folderWebUrl} target="_blank" rel="noopener noreferrer" style={{ fontSize: 12, color: theme.palette.themePrimary }}>
                Open folder in SharePoint
              </a>
            )}
          </Stack>
        )}
      </div>

      {isLoading && <Spinner size={SpinnerSize.small} label={`Calling ${state.loading}...`} styles={{ root: { marginTop: 8 } }} />}
    </div>
  );
};
