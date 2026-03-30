/**
 * Outlook Mail MCP Server catalog — server ID and examples only.
 * Tool definitions come from live discovery.
 */

export interface IToolExample {
  label: string;
  params: Record<string, unknown>;
}

export const MAIL_SERVER_ID = 'mcp_MailTools';

export const MAIL_SERVER_EXAMPLES: Record<string, IToolExample[]> = {
  SearchMessages: [
    { label: 'Search inbox', params: { message: 'project update' } },
    { label: 'From specific sender', params: { message: 'from:boss@contoso.com' } }
  ],
  GetMessage: [
    { label: 'Get by ID', params: { messageId: 'message-id-here' } }
  ],
  CreateDraftMessage: [
    { label: 'Simple draft', params: { subject: 'Test from MCP Explorer', bodyContent: 'This is a test email sent via the MCP Mail server.', toRecipients: 'colleague@contoso.com' } }
  ],
  SendEmailWithAttachments: [
    { label: 'Quick email', params: { subject: 'Hello from MCP Explorer', bodyContent: 'Testing the Work IQ Mail server from SPFx.', toRecipients: 'colleague@contoso.com' } }
  ],
  ReplyToMessage: [
    { label: 'Reply', params: { messageId: 'message-id-here', comment: 'Thanks for the update!' } }
  ]
};
