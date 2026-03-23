/**
 * SharePoint Lists MCP Server catalog — hardcoded fallback from Grimoire discovery.
 * Live discovery from the server is preferred when connected.
 */

export interface IMcpCatalogTool {
  name: string;
  description: string;
  inputSchema: {
    type: string;
    properties: Record<string, { type: string; description: string; items?: { type: string } }>;
    required?: string[];
  };
}

export const LISTS_SERVER_ID = 'mcp_SharePointRemoteServer';
export const LISTS_SERVER_SCOPE = 'McpServers.SharepointLists.All';

export const LISTS_SERVER_TOOLS: IMcpCatalogTool[] = [
  { name: 'searchSitesByName', description: 'Find SharePoint sites by name, title, or partial name.', inputSchema: { type: 'object', properties: { search: { type: 'string', description: 'Display name or partial name of the site to search for.' }, consistencyLevel: { type: 'string', description: 'Required by Microsoft Graph for search queries.' } }, required: ['search'] } },
  { name: 'getSiteByPath', description: 'Resolve a SharePoint site using its exact hostname and server-relative path.', inputSchema: { type: 'object', properties: { hostname: { type: 'string', description: 'Host name of the SharePoint tenant (e.g., contoso.sharepoint.com).' }, serverRelativePath: { type: 'string', description: 'Server-relative path to the site (e.g., sites/Marketing).' } }, required: ['hostname', 'serverRelativePath'] } },
  { name: 'listLists', description: 'Get all SharePoint lists available on a specific site.', inputSchema: { type: 'object', properties: { siteId: { type: 'string', description: 'The full site ID in format hostname,siteCollectionId,webId.' } }, required: ['siteId'] } },
  { name: 'listSubsites', description: 'List child sites (subsites) for a given site.', inputSchema: { type: 'object', properties: { siteId: { type: 'string', description: 'The unique ID of the site.' } }, required: ['siteId'] } },
  { name: 'listListColumns', description: 'List column definitions for a specific SharePoint list.', inputSchema: { type: 'object', properties: { siteId: { type: 'string', description: 'The unique ID of the site.' }, listId: { type: 'string', description: 'The unique ID of the list.' }, select: { type: 'string', description: 'Comma-separated list of properties to return.' }, filter: { type: 'string', description: 'OData filter expression.' }, orderBy: { type: 'string', description: 'Property to order by.' }, top: { type: 'string', description: 'Page size.' } }, required: ['siteId', 'listId'] } },
  { name: 'listListItems', description: 'Get items (rows/records) from a specific SharePoint list.', inputSchema: { type: 'object', properties: { siteId: { type: 'string', description: 'The full site ID.' }, listId: { type: 'string', description: 'The unique ID of the list.' }, expand: { type: 'string', description: 'OData $expand clause. Defaults to "fields".' }, top: { type: 'string', description: 'Page size.' }, filter: { type: 'string', description: 'OData filter expression.' }, select: { type: 'string', description: 'OData select clause.' } }, required: ['siteId', 'listId'] } },
  { name: 'createList', description: 'Create a new SharePoint list on a site.', inputSchema: { type: 'object', properties: { siteId: { type: 'string', description: 'The full site ID.' }, displayName: { type: 'string', description: 'Display name of the list.' }, list: { type: 'string', description: 'List info such as template (JSON object).' } }, required: ['siteId', 'displayName', 'list'] } },
  { name: 'createListColumn', description: 'Create a new column in a specific SharePoint list.', inputSchema: { type: 'object', properties: { siteId: { type: 'string', description: 'The unique ID of the site.' }, listId: { type: 'string', description: 'The unique ID of the list.' }, name: { type: 'string', description: 'API/static name of the column (no spaces).' }, displayName: { type: 'string', description: 'User-facing display name.' }, description: { type: 'string', description: 'Column description.' } }, required: ['siteId', 'listId', 'name'] } },
  { name: 'createListItem', description: 'Create a new item (row/record) in a specific SharePoint list.', inputSchema: { type: 'object', properties: { siteId: { type: 'string', description: 'The full site ID.' }, listId: { type: 'string', description: 'The unique ID of the list.' }, fields: { type: 'string', description: 'Key-value pairs of column names and values (JSON object).' } }, required: ['siteId', 'listId', 'fields'] } },
  { name: 'updateListItem', description: 'Update fields of an existing item in a SharePoint list.', inputSchema: { type: 'object', properties: { siteId: { type: 'string', description: 'The unique ID of the site.' }, listId: { type: 'string', description: 'The unique ID of the list.' }, itemId: { type: 'string', description: 'The unique ID of the list item.' }, fields: { type: 'string', description: 'Key-value pairs of fields to update (JSON object).' }, ifMatch: { type: 'string', description: 'ETag for concurrency control. Use "*" to force update.' } }, required: ['siteId', 'listId', 'itemId', 'fields'] } },
  { name: 'editListColumn', description: 'Update an existing column on a specific SharePoint list.', inputSchema: { type: 'object', properties: { siteId: { type: 'string', description: 'Unique ID of the site.' }, listId: { type: 'string', description: 'Unique ID of the list.' }, columnId: { type: 'string', description: 'Unique ID of the column to update.' }, displayName: { type: 'string', description: 'New display name.' }, description: { type: 'string', description: 'New description.' } }, required: ['siteId', 'listId', 'columnId'] } },
  { name: 'deleteListItem', description: 'Delete an item (row/record) from a SharePoint list.', inputSchema: { type: 'object', properties: { siteId: { type: 'string', description: 'The full site ID.' }, listId: { type: 'string', description: 'The unique ID of the list.' }, itemId: { type: 'string', description: 'The unique ID of the list item.' }, ifMatch: { type: 'string', description: 'ETag for concurrency control.' } }, required: ['siteId', 'listId', 'itemId'] } },
  { name: 'deleteListColumn', description: 'Delete an existing column from a specific SharePoint list.', inputSchema: { type: 'object', properties: { siteId: { type: 'string', description: 'Unique ID of the site.' }, listId: { type: 'string', description: 'Unique ID of the list.' }, columnId: { type: 'string', description: 'Unique ID of the column to delete.' } }, required: ['siteId', 'listId', 'columnId'] } }
];

export interface IToolExample {
  label: string;
  params: Record<string, unknown>;
}

export const LISTS_SERVER_EXAMPLES: Record<string, IToolExample[]> = {
  searchSitesByName: [
    { label: 'Search by name', params: { search: 'Marketing', consistencyLevel: 'eventual' } }
  ],
  getSiteByPath: [
    { label: 'Site by path', params: { hostname: 'contoso.sharepoint.com', serverRelativePath: 'sites/Marketing' } }
  ],
  listLists: [
    { label: 'All lists on site', params: { siteId: 'contoso.sharepoint.com,guid1,guid2' } }
  ],
  listListItems: [
    { label: 'First 10 items', params: { siteId: 'contoso.sharepoint.com,guid1,guid2', listId: 'list-guid', top: '10' } }
  ],
  listListColumns: [
    { label: 'All columns', params: { siteId: 'contoso.sharepoint.com,guid1,guid2', listId: 'list-guid' } }
  ]
};
