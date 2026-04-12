/**
 * Microsoft OneDrive MCP Server catalog — server ID and examples.
 * Tool list discovered live via tools/list on 2026-04-11.
 *
 * mcp_OneDriveRemoteServer is the new dedicated OneDrive server (replaces the
 * deprecated mcp_ODSPRemoteServer combined SharePoint+OneDrive server).
 * 13 tools, all scoped to the user's personal OneDrive.
 */

export interface IToolExample {
  label: string;
  params: Record<string, unknown>;
}

export const ONEDRIVE_SERVER_ID = 'mcp_OneDriveRemoteServer';

export const ONEDRIVE_SERVER_EXAMPLES: Record<string, IToolExample[]> = {
  getOnedrive: [
    { label: 'My OneDrive overview', params: {} }
  ],
  findFileOrFolderInMyDrive: [
    { label: 'Search by name', params: { searchQuery: 'report' } }
  ],
  getFolderChildrenInMyOnedrive: [
    { label: 'Root folder', params: {} },
    { label: 'Specific folder', params: { parentFolderId: 'folder-id-here' } }
  ],
  getFileOrFolderMetadataInMyOnedrive: [
    { label: 'By item ID', params: { fileOrFolderId: 'item-id-here' } }
  ],
  getFileOrFolderMetadataByUrl: [
    { label: 'By sharing URL', params: { fileOrFolderUrl: 'https://contoso-my.sharepoint.com/personal/.../Documents/file.docx' } }
  ],
  readSmallTextFileFromMyOnedrive: [
    { label: 'Read a text file', params: { fileId: 'file-id-here' } }
  ],
  createSmallTextFileInMyOnedrive: [
    { label: 'Create hello.txt in root', params: { fileName: 'hello.txt', content: 'Hello from MCP365 Explorer' } }
  ],
  createFolderInMyOnedrive: [
    { label: 'Create folder in root', params: { folderName: 'MCP365_Demo' } }
  ],
  renameFileOrFolderInMyOnedrive: [
    { label: 'Rename an item', params: { fileOrFolderId: 'item-id-here', newName: 'new-name.txt' } }
  ],
  moveSmallFileInMyOnedrive: [
    { label: 'Move file to a folder', params: { fileId: 'file-id-here', newParentFolderId: 'folder-id-here' } }
  ],
  deleteFileOrFolderInMyOnedrive: [
    { label: 'Delete an item', params: { fileOrFolderId: 'item-id-here' } }
  ],
  shareFileOrFolderInMyOnedrive: [
    { label: 'Share with a user', params: { fileOrFolderId: 'item-id-here', recipients: ['colleague@contoso.com'], role: 'read' } }
  ],
  setSensitivityLabelOnFileInMyOnedrive: [
    { label: 'Apply a label', params: { fileId: 'file-id-here', labelId: 'label-id-here' } }
  ]
};
