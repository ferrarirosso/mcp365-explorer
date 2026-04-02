/**
 * Microsoft Teams MCP Server catalog — server ID and examples only.
 * Tool definitions come from live discovery.
 */

export interface IToolExample {
  label: string;
  params: Record<string, unknown>;
}

export const TEAMS_SERVER_ID = 'mcp_TeamsServer';

export const TEAMS_SERVER_EXAMPLES: Record<string, IToolExample[]> = {
  ListTeams: [
    { label: 'My teams', params: {} }
  ],
  ListChats: [
    { label: 'My chats', params: {} }
  ],
  ListChannels: [
    { label: 'Channels in a team', params: { teamId: 'team-id-here' } }
  ],
  SearchTeamsMessages: [
    { label: 'Search for keyword', params: { searchQuery: 'project update' } }
  ],
  PostMessage: [
    { label: 'Send chat message', params: { chatId: 'chat-id-here', message: 'Hello from MCP Explorer!' } }
  ],
  CreateChat: [
    { label: 'New 1:1 chat', params: { chatType: 'oneOnOne', members: ['colleague@contoso.com'] } }
  ]
};
