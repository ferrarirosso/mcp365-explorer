/**
 * User Profile MCP Server catalog — hardcoded from Grimoire discovery output.
 * Used as fallback reference when not connected, and for example queries.
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

export const ME_SERVER_ID = 'mcp_MeServer';
export const ME_SERVER_SCOPE = 'McpServers.Me.All';

export const ME_SERVER_TOOLS: IMcpCatalogTool[] = [
  {
    name: 'GetMyDetails',
    description: 'Retrieve profile details for the currently signed-in user ("me, my"). Use this when you need the signed-in user\'s identity or profile details (e.g., display name, email, job title).',
    inputSchema: {
      type: 'object',
      properties: {
        select: { type: 'string', description: 'Comma-separated list of properties you need' },
        expand: { type: 'string', description: 'Expand related entities' }
      }
    }
  },
  {
    name: 'GetUserDetails',
    description: 'Find a specified user\'s profile by name, email, or ID. Use this when you need to look up a specific person in your organization.',
    inputSchema: {
      type: 'object',
      properties: {
        userIdentifier: { type: 'string', description: 'The user\'s name, object ID (GUID), or userPrincipalName (email-like UPN).' },
        select: { type: 'string', description: 'Comma-separated list of properties you need' },
        expand: { type: 'string', description: 'Expand a related entity for the user' }
      },
      required: ['userIdentifier']
    }
  },
  {
    name: 'GetMultipleUsersDetails',
    description: 'Search for multiple users in the directory by name, job title, office location, or other properties.',
    inputSchema: {
      type: 'object',
      properties: {
        searchValues: { type: 'array', description: 'List of search terms (e.g., ["John Smith", "Jane Doe"])', items: { type: 'string' } },
        propertyToSearchBy: { type: 'string', description: 'User property to search (e.g., "displayName", "jobTitle", "officeLocation")' },
        select: { type: 'string', description: 'Comma-separated list of user properties to include in response' },
        expand: { type: 'string', description: 'Navigation properties to expand (e.g., "manager")' },
        top: { type: 'string', description: 'Maximum number of results to return' },
        orderby: { type: 'string', description: 'Property name to sort results by' }
      },
      required: ['searchValues']
    }
  },
  {
    name: 'GetManagerDetails',
    description: 'Get a user\'s manager information — name, email, job title, etc.',
    inputSchema: {
      type: 'object',
      properties: {
        userId: { type: 'string', description: 'Name of the user whose manager to retrieve. Use "me" for current user.' },
        select: { type: 'string', description: 'Comma-separated list of properties you need' }
      },
      required: ['userId']
    }
  },
  {
    name: 'GetDirectReportsDetails',
    description: 'Retrieve a user\'s direct reports (people who report to them in the org hierarchy). Use for organizational team structure, NOT for Microsoft Teams membership.',
    inputSchema: {
      type: 'object',
      properties: {
        userId: { type: 'string', description: 'Name of the user whose direct reports to retrieve. Use "me" for current user.' },
        select: { type: 'string', description: 'Comma-separated list of properties you need for each direct report' }
      },
      required: ['userId']
    }
  }
];

export interface IToolExample {
  label: string;
  params: Record<string, unknown>;
}

export const ME_SERVER_EXAMPLES: Record<string, IToolExample[]> = {
  GetMyDetails: [
    { label: 'Basic profile', params: { select: 'displayName,mail,jobTitle' } },
    { label: 'Full profile with manager', params: { select: 'displayName,mail,jobTitle,department,officeLocation', expand: 'manager' } }
  ],
  GetUserDetails: [
    { label: 'Lookup by email', params: { userIdentifier: 'user@contoso.com', select: 'displayName,mail,jobTitle' } },
    { label: 'Lookup by name', params: { userIdentifier: 'Adele Vance', select: 'displayName,mail,department' } }
  ],
  GetMultipleUsersDetails: [
    { label: 'Search by job title', params: { searchValues: ['Software Engineer', 'Product Manager'], propertyToSearchBy: 'jobTitle', select: 'displayName,mail,jobTitle' } },
    { label: 'Search by name', params: { searchValues: ['John', 'Jane'], select: 'displayName,mail,jobTitle,department' } }
  ],
  GetManagerDetails: [
    { label: 'My manager', params: { userId: 'me', select: 'displayName,mail,jobTitle' } }
  ],
  GetDirectReportsDetails: [
    { label: 'My direct reports', params: { userId: 'me', select: 'displayName,mail,jobTitle,department' } }
  ]
};
