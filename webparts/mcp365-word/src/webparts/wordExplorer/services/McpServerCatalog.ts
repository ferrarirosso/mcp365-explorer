/**
 * Microsoft Word MCP Server catalog — server ID and examples only.
 * Tool definitions come from live discovery.
 *
 * mcp_WordServer (4 tools):
 *   - CreateDocument(fileName, contentInHtml, shareWith?)
 *   - GetDocumentContent(url) → { filename, size, driveId, documentId, content, comments }
 *   - AddComment(driveId, documentId, newComment)
 *   - ReplyToComment(commentId, driveId, documentId, newComment)
 */

export interface IToolExample {
  label: string;
  params: Record<string, unknown>;
}

export const WORD_SERVER_ID = 'mcp_WordServer';

export const WORD_SERVER_EXAMPLES: Record<string, IToolExample[]> = {
  CreateDocument: [
    {
      label: 'Hello world document',
      params: {
        fileName: 'MCP365_Hello.docx',
        contentInHtml: '<h1>Hello from MCP365 Explorer</h1><p>Created via <code>mcp_WordServer.CreateDocument</code>.</p>'
      }
    },
    {
      label: 'Empty document (auto-named)',
      params: {
        fileName: '',
        contentInHtml: '<p>Draft document.</p>'
      }
    }
  ],
  GetDocumentContent: [
    { label: 'Read a Word file by URL', params: { url: 'https://contoso.sharepoint.com/personal/.../Documents/sample.docx' } }
  ],
  AddComment: [
    { label: 'Add a comment', params: { driveId: 'drive-id-here', documentId: 'document-id-here', newComment: 'Looks good to me.' } }
  ],
  ReplyToComment: [
    { label: 'Reply to a comment', params: { commentId: 'comment-id-here', driveId: 'drive-id-here', documentId: 'document-id-here', newComment: 'Thanks!' } }
  ]
};
