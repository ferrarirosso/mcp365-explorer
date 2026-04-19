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

interface ICreatedDoc {
  name: string;
  webUrl: string;
  driveId: string | undefined;
  documentId: string | undefined;
}

interface IDocContent {
  filename: string;
  size: number | undefined;
  driveId: string;
  documentId: string;
  contentExcerpt: string;
  commentCount: number;
}

interface IShowcaseState {
  // Step 1 — CreateDocument
  fileName: string;
  htmlContent: string;
  createdDoc: ICreatedDoc | undefined;

  // Step 2 — GetDocumentContent
  readUrl: string;
  docContent: IDocContent | undefined;

  // Step 3 — AddComment
  commentText: string;
  addedCommentId: string | undefined;

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
    if (outer && typeof outer === 'object' && typeof (outer as Record<string, unknown>).response === 'string') {
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

/**
 * CreateDocument wraps the DriveItem in { driveItem: {...} }. If we see that
 * wrapper, return the inner object; otherwise return obj as-is.
 */
function unwrapDriveItem(obj: Record<string, unknown> | undefined): Record<string, unknown> | undefined {
  if (!obj) return undefined;
  if (obj.driveItem && typeof obj.driveItem === 'object' && !Array.isArray(obj.driveItem)) {
    return obj.driveItem as Record<string, unknown>;
  }
  return obj;
}

/** Case-insensitive string getter — tries camelCase key first, then PascalCase. */
function getStringCI(obj: Record<string, unknown> | undefined, key: string): string | undefined {
  if (!obj) return undefined;
  const v1 = obj[key];
  if (typeof v1 === 'string') return v1;
  const pascal = key.charAt(0).toUpperCase() + key.slice(1);
  const v2 = obj[pascal];
  if (typeof v2 === 'string') return v2;
  return undefined;
}

/** Case-insensitive number getter. */
function getNumberCI(obj: Record<string, unknown> | undefined, key: string): number | undefined {
  if (!obj) return undefined;
  const v1 = obj[key];
  if (typeof v1 === 'number') return v1;
  const pascal = key.charAt(0).toUpperCase() + key.slice(1);
  const v2 = obj[pascal];
  if (typeof v2 === 'number') return v2;
  return undefined;
}

/** Case-insensitive nested-object getter. */
function getObjectCI(obj: Record<string, unknown> | undefined, key: string): Record<string, unknown> | undefined {
  if (!obj) return undefined;
  const v1 = obj[key];
  if (v1 && typeof v1 === 'object' && !Array.isArray(v1)) return v1 as Record<string, unknown>;
  const pascal = key.charAt(0).toUpperCase() + key.slice(1);
  const v2 = obj[pascal];
  if (v2 && typeof v2 === 'object' && !Array.isArray(v2)) return v2 as Record<string, unknown>;
  return undefined;
}

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

/** Try multiple keys to find a URL in a DriveItem response (handles camelCase and PascalCase). */
function extractUrl(obj: Record<string, unknown> | undefined): string {
  if (!obj) return '';
  const webUrl = getStringCI(obj, 'webUrl');
  if (webUrl) return webUrl;
  const dlUrl = getString(obj, '@microsoft.graph.downloadUrl');
  if (dlUrl) return dlUrl;
  const spIds = getObjectCI(obj, 'sharepointIds');
  const siteUrl = getStringCI(spIds, 'siteUrl');
  if (siteUrl) return siteUrl;
  // Scan all string values for a SharePoint URL (last resort)
  const keys = Object.keys(obj);
  for (let i = 0; i < keys.length; i++) {
    const v = obj[keys[i]];
    if (typeof v === 'string' && v.startsWith('https://') && v.indexOf('sharepoint.com') > -1) return v;
  }
  return '';
}

function getString(obj: Record<string, unknown> | undefined, key: string): string | undefined {
  if (!obj) return undefined;
  const v = obj[key];
  return typeof v === 'string' ? v : undefined;
}

const DEFAULT_HTML = '<h1>Hello from MCP365 Explorer</h1><p>This document was created by calling <code>mcp_WordServer.CreateDocument</code> over JSON-RPC, directly from an SPFx web part — no Graph SDK, no orchestration layer.</p><p>The next step in the showcase reads it back via <code>GetDocumentContent</code>, then adds a comment with <code>AddComment</code>.</p>';

export const WordShowcase: React.FC<IShowcaseProps> = ({ client, theme }) => {
  const [state, setState] = React.useState<IShowcaseState>({
    fileName: `MCP365_Demo_${new Date().toISOString().slice(0, 10).replace(/-/g, '')}.docx`,
    htmlContent: DEFAULT_HTML,
    createdDoc: undefined,
    readUrl: '',
    docContent: undefined,
    commentText: 'Looks great — generated end-to-end via MCP.',
    addedCommentId: undefined,
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

  // ── Step 1: Create the document ───────────────────────────

  const handleCreate = React.useCallback(async (): Promise<void> => {
    if (!state.htmlContent.trim()) return;
    setState(prev => ({ ...prev, loading: 'CreateDocument', error: undefined, createdDoc: undefined }));
    try {
      const result = await client.callTool('CreateDocument', {
        fileName: state.fileName,
        contentInHtml: state.htmlContent
      });
      const err = extractError(result);
      if (err) {
        setState(prev => ({ ...prev, loading: undefined, error: err }));
        return;
      }
      // CreateDocument wraps the DriveItem in { driveItem: {...} } with PascalCase keys
      const item = unwrapDriveItem(extractObject(result));
      const webUrl = extractUrl(item);
      const name = getStringCI(item, 'name') || state.fileName || 'Document';
      const documentId = getStringCI(item, 'id');
      const parent = getObjectCI(item, 'parentReference');
      const driveId = getStringCI(parent, 'driveId');
      setState(prev => ({
        ...prev,
        loading: undefined,
        createdDoc: { name, webUrl, driveId, documentId },
        readUrl: webUrl  // auto-populate step 2
      }));
    } catch (err) { setState(prev => ({ ...prev, loading: undefined, error: (err as Error).message })); }
  }, [client, state.fileName, state.htmlContent]);

  // ── Step 2: Read the document content ──────────────────────

  const handleRead = React.useCallback(async (): Promise<void> => {
    if (!state.readUrl.trim()) return;
    setState(prev => ({ ...prev, loading: 'GetDocumentContent', error: undefined, docContent: undefined }));
    try {
      const result = await client.callTool('GetDocumentContent', { url: state.readUrl.trim() });
      const err = extractError(result);
      if (err) {
        setState(prev => ({ ...prev, loading: undefined, error: err }));
        return;
      }
      const obj = extractObject(result);
      const filename = getStringCI(obj, 'filename') || getStringCI(obj, 'name') || '(unknown)';
      const size = getNumberCI(obj, 'size');
      const driveId = getStringCI(obj, 'driveId') || '';
      const documentId = getStringCI(obj, 'documentId') || '';
      const content = getStringCI(obj, 'content') || '';
      const commentsVal = obj ? (obj.comments || obj.Comments) : undefined;
      const comments = Array.isArray(commentsVal) ? commentsVal : [];
      const excerpt = content.length > 240 ? content.substring(0, 240) + '…' : content;
      setState(prev => ({
        ...prev,
        loading: undefined,
        docContent: { filename, size, driveId, documentId, contentExcerpt: excerpt, commentCount: comments.length }
      }));
    } catch (err) { setState(prev => ({ ...prev, loading: undefined, error: (err as Error).message })); }
  }, [client, state.readUrl]);

  // ── Step 3: Add a comment ──────────────────────────────────

  const handleAddComment = React.useCallback(async (): Promise<void> => {
    if (!state.docContent) return;
    if (!state.commentText.trim()) return;
    setState(prev => ({ ...prev, loading: 'AddComment', error: undefined, addedCommentId: undefined }));
    try {
      const result = await client.callTool('AddComment', {
        driveId: state.docContent.driveId,
        documentId: state.docContent.documentId,
        newComment: state.commentText.trim()
      });
      const err = extractError(result);
      if (err) {
        setState(prev => ({ ...prev, loading: undefined, error: err }));
        return;
      }
      const obj = extractObject(result);
      const id = getStringCI(obj, 'id') || getStringCI(obj, 'commentId') || 'added';
      setState(prev => ({ ...prev, loading: undefined, addedCommentId: id }));
    } catch (err) { setState(prev => ({ ...prev, loading: undefined, error: (err as Error).message })); }
  }, [client, state.docContent, state.commentText]);

  // ── Render ────────────────────────────────────────────────

  return (
    <div style={{ marginTop: 12 }}>
      <Text variant="small" styles={{ root: { color: theme.palette.neutralTertiary, marginBottom: 12, display: 'block' } }}>
        Create a Word document, read it back, and add a comment — three chained MCP tool calls.
      </Text>

      {state.error && (
        <Text variant="small" styles={{ root: { color: theme.palette.red, marginBottom: 8, display: 'block' } }}>{state.error}</Text>
      )}

      {/* ── Step 1 ─────────────────────────────────────────────── */}
      <div style={cardStyle}>
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }} styles={{ root: { marginBottom: 8 } }}>
          <span style={stepBadgeStyle}>1</span>
          <Text styles={{ root: { fontWeight: 600 } }}>Create a document</Text>
          <Text variant="small" styles={{ root: { color: theme.palette.neutralTertiary } }}>(CreateDocument)</Text>
        </Stack>
        <Stack tokens={{ childrenGap: 8 }}>
          <TextField
            label="File name"
            value={state.fileName}
            onChange={(_, val) => setState(prev => ({ ...prev, fileName: val || '' }))}
            disabled={isLoading}
          />
          <TextField
            label="Body (HTML)"
            multiline
            rows={4}
            value={state.htmlContent}
            onChange={(_, val) => setState(prev => ({ ...prev, htmlContent: val || '' }))}
            disabled={isLoading}
          />
          <PrimaryButton
            text={state.loading === 'CreateDocument' ? 'Creating...' : 'Create document'}
            iconProps={{ iconName: 'WordDocument' }}
            onClick={handleCreate}
            disabled={isLoading || !state.htmlContent.trim()}
            styles={{ root: { alignSelf: 'flex-start' } }}
          />
        </Stack>

        {state.createdDoc && (
          <div style={{ marginTop: 8, padding: 8, backgroundColor: theme.palette.neutralLighterAlt, borderRadius: 4 }}>
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 6 }}>
              <Icon iconName="CheckMark" styles={{ root: { color: '#107c10' } }} />
              <Text variant="small" styles={{ root: { fontWeight: 600 } }}>{state.createdDoc.name}</Text>
            </Stack>
            {state.createdDoc.webUrl && (
              <Text variant="small" styles={{ root: { display: 'block', marginTop: 4, wordBreak: 'break-all' } }}>
                <a href={state.createdDoc.webUrl} target="_blank" rel="noreferrer" data-interception="off">{state.createdDoc.webUrl}</a>
              </Text>
            )}
          </div>
        )}
      </div>

      {/* ── Step 2 ─────────────────────────────────────────────── */}
      <div style={cardStyle}>
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }} styles={{ root: { marginBottom: 8 } }}>
          <span style={stepBadgeStyle}>2</span>
          <Text styles={{ root: { fontWeight: 600 } }}>Read the document</Text>
          <Text variant="small" styles={{ root: { color: theme.palette.neutralTertiary } }}>(GetDocumentContent)</Text>
        </Stack>
        <Stack tokens={{ childrenGap: 8 }}>
          <TextField
            label="Document URL"
            placeholder="Auto-filled after step 1, or paste any OneDrive/SharePoint .docx URL"
            value={state.readUrl}
            onChange={(_, val) => setState(prev => ({ ...prev, readUrl: val || '' }))}
            disabled={isLoading}
          />
          <DefaultButton
            text={state.loading === 'GetDocumentContent' ? 'Reading...' : 'Read document'}
            iconProps={{ iconName: 'OpenFile' }}
            onClick={handleRead}
            disabled={isLoading || !state.readUrl.trim()}
            styles={{ root: { alignSelf: 'flex-start' } }}
          />
        </Stack>

        {state.docContent && (
          <div style={{ marginTop: 8, padding: 8, backgroundColor: theme.palette.neutralLighterAlt, borderRadius: 4 }}>
            <Text variant="small" styles={{ root: { fontWeight: 600, display: 'block' } }}>
              {state.docContent.filename}
              {state.docContent.size !== undefined && ` — ${state.docContent.size} bytes`}
              {` — ${state.docContent.commentCount} comment${state.docContent.commentCount !== 1 ? 's' : ''}`}
            </Text>
            <Text variant="small" styles={{ root: { color: theme.palette.neutralSecondary, marginTop: 4, display: 'block', whiteSpace: 'pre-wrap' } }}>
              {state.docContent.contentExcerpt || '(empty)'}
            </Text>
            <Text variant="small" styles={{ root: { color: theme.palette.neutralTertiary, marginTop: 4, display: 'block', fontFamily: 'Consolas, monospace' } }}>
              driveId: {state.docContent.driveId.substring(0, 30)}…
              {' · '}
              documentId: {state.docContent.documentId.substring(0, 20)}…
            </Text>
          </div>
        )}
      </div>

      {/* ── Step 3 ─────────────────────────────────────────────── */}
      <div style={cardStyle}>
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }} styles={{ root: { marginBottom: 8 } }}>
          <span style={stepBadgeStyle}>3</span>
          <Text styles={{ root: { fontWeight: 600 } }}>Add a comment</Text>
          <Text variant="small" styles={{ root: { color: theme.palette.neutralTertiary } }}>(AddComment — needs driveId + documentId from step 2)</Text>
        </Stack>
        <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="end">
          <TextField
            placeholder="Comment text"
            value={state.commentText}
            onChange={(_, val) => setState(prev => ({ ...prev, commentText: val || '' }))}
            disabled={isLoading || !state.docContent}
            styles={{ root: { width: 380 } }}
            onKeyDown={(e) => { if (e.key === 'Enter') { void handleAddComment(); } }}
          />
          <DefaultButton
            text={state.loading === 'AddComment' ? 'Adding...' : 'Add comment'}
            iconProps={{ iconName: 'Comment' }}
            onClick={handleAddComment}
            disabled={isLoading || !state.docContent || !state.commentText.trim()}
          />
        </Stack>
        {!state.docContent && (
          <Text variant="small" styles={{ root: { color: theme.palette.neutralTertiary, marginTop: 6, display: 'block' } }}>
            Run step 2 first to capture driveId and documentId.
          </Text>
        )}
        {state.addedCommentId && (
          <Text variant="small" styles={{ root: { color: '#107c10', fontWeight: 600, marginTop: 6, display: 'block' } }}>
            Comment added (id: {state.addedCommentId}). Open the document in Word to see it.
          </Text>
        )}
      </div>

      {isLoading && <Spinner size={SpinnerSize.small} label={`Calling ${state.loading}...`} styles={{ root: { marginTop: 8 } }} />}
    </div>
  );
};
