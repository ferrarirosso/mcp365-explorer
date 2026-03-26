/**
 * Outlook Calendar MCP Server catalog — hardcoded fallback from Grimoire discovery.
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

export const CALENDAR_SERVER_ID = 'mcp_CalendarTools';
export const CALENDAR_SERVER_SCOPE = 'McpServers.Calendar.All';

export const CALENDAR_SERVER_TOOLS: IMcpCatalogTool[] = [
  { name: 'ListEvents', description: 'Retrieve a list of events for the user with given criteria.', inputSchema: { type: 'object', properties: { startDateTime: { type: 'string', description: 'Start of the time range (ISO 8601).' }, endDateTime: { type: 'string', description: 'End of the time range (ISO 8601).' }, meetingTitle: { type: 'string', description: 'Filter by meeting title.' }, attendeeEmails: { type: 'string', description: 'Filter by attendee email.' }, timeZone: { type: 'string', description: 'Time zone (e.g., "Europe/Zurich").' }, select: { type: 'string', description: 'Comma-separated properties to return.' }, top: { type: 'string', description: 'Max results.' }, orderby: { type: 'string', description: 'Order by property.' } } } },
  { name: 'ListCalendarView', description: 'Retrieve events from a calendar view with recurring events expanded.', inputSchema: { type: 'object', properties: { userIdentifier: { type: 'string', description: 'User identifier or "me".' }, startDateTime: { type: 'string', description: 'Start of the time range (ISO 8601).' }, endDateTime: { type: 'string', description: 'End of the time range (ISO 8601).' }, timeZone: { type: 'string', description: 'Time zone.' }, subject: { type: 'string', description: 'Filter by subject.' }, select: { type: 'string', description: 'Properties to return.' }, top: { type: 'string', description: 'Max results.' }, orderby: { type: 'string', description: 'Order by property.' } }, required: ['userIdentifier'] } },
  { name: 'CreateEvent', description: 'Create a new calendar event.', inputSchema: { type: 'object', properties: { subject: { type: 'string', description: 'Event subject/title.' }, attendeeEmails: { type: 'array', description: 'List of attendee email addresses.', items: { type: 'string' } }, startDateTime: { type: 'string', description: 'Start date/time (ISO 8601).' }, endDateTime: { type: 'string', description: 'End date/time (ISO 8601).' }, timeZone: { type: 'string', description: 'Time zone.' }, bodyContent: { type: 'string', description: 'Event body content.' }, location: { type: 'string', description: 'Event location.' }, isOnlineMeeting: { type: 'string', description: 'Set to "true" for Teams meeting.' } }, required: ['subject', 'attendeeEmails', 'startDateTime', 'endDateTime'] } },
  { name: 'UpdateEvent', description: 'Update an existing calendar event.', inputSchema: { type: 'object', properties: { eventId: { type: 'string', description: 'The ID of the event to update.' }, subject: { type: 'string', description: 'New subject.' }, startDateTime: { type: 'string', description: 'New start date/time.' }, endDateTime: { type: 'string', description: 'New end date/time.' }, timeZone: { type: 'string', description: 'Time zone.' }, location: { type: 'string', description: 'New location.' } }, required: ['eventId'] } },
  { name: 'DeleteEventById', description: 'Delete a calendar event.', inputSchema: { type: 'object', properties: { eventId: { type: 'string', description: 'The ID of the event to delete.' } }, required: ['eventId'] } },
  { name: 'FindMeetingTimes', description: 'Find meeting times that work for all attendees.', inputSchema: { type: 'object', properties: { userIdentifier: { type: 'string', description: 'Organizer identifier or "me".' }, attendeeEmails: { type: 'array', description: 'List of attendee email addresses.', items: { type: 'string' } }, meetingDuration: { type: 'string', description: 'Duration in ISO 8601 format (e.g., "PT1H").' }, startDateTime: { type: 'string', description: 'Start of search range.' }, endDateTime: { type: 'string', description: 'End of search range.' }, timeZone: { type: 'string', description: 'Time zone.' }, maxCandidates: { type: 'string', description: 'Max number of suggestions.' } }, required: ['userIdentifier', 'attendeeEmails', 'meetingDuration'] } },
  { name: 'AcceptEvent', description: 'Accept a calendar event invitation.', inputSchema: { type: 'object', properties: { eventId: { type: 'string', description: 'The event ID.' }, comment: { type: 'string', description: 'Optional comment.' }, sendResponse: { type: 'string', description: 'Whether to send a response ("true"/"false").' } }, required: ['eventId'] } },
  { name: 'TentativelyAcceptEvent', description: 'Tentatively accept a calendar event invitation.', inputSchema: { type: 'object', properties: { eventId: { type: 'string', description: 'The event ID.' }, comment: { type: 'string', description: 'Optional comment.' }, sendResponse: { type: 'string', description: 'Whether to send a response.' } }, required: ['eventId'] } },
  { name: 'DeclineEvent', description: 'Decline a calendar event invitation.', inputSchema: { type: 'object', properties: { eventId: { type: 'string', description: 'The event ID.' }, comment: { type: 'string', description: 'Optional comment.' }, sendResponse: { type: 'string', description: 'Whether to send a response.' } }, required: ['eventId'] } },
  { name: 'CancelEvent', description: 'Cancel a calendar event (organizer only).', inputSchema: { type: 'object', properties: { eventId: { type: 'string', description: 'The event ID.' }, comment: { type: 'string', description: 'Optional cancellation message.' } }, required: ['eventId'] } },
  { name: 'ForwardEvent', description: 'Forward a calendar event to other recipients.', inputSchema: { type: 'object', properties: { eventId: { type: 'string', description: 'The event ID.' }, recipientEmails: { type: 'array', description: 'List of recipient email addresses.', items: { type: 'string' } }, comment: { type: 'string', description: 'Optional comment.' } }, required: ['eventId', 'recipientEmails'] } },
  { name: 'GetUserDateAndTimeZoneSettings', description: 'Get date and timezone settings for a user.', inputSchema: { type: 'object', properties: { userIdentifier: { type: 'string', description: 'User identifier or "me".' } } } },
  { name: 'GetRooms', description: 'Get all meeting rooms defined in the tenant.', inputSchema: { type: 'object', properties: {} } }
];

export interface IToolExample {
  label: string;
  params: Record<string, unknown>;
}

export const CALENDAR_SERVER_EXAMPLES: Record<string, IToolExample[]> = {
  ListEvents: [
    { label: 'My events this week', params: { startDateTime: '2026-03-18T00:00:00', endDateTime: '2026-03-24T23:59:59', timeZone: 'Europe/Zurich' } },
    { label: 'Today only', params: { startDateTime: '2026-03-18T00:00:00', endDateTime: '2026-03-18T23:59:59', timeZone: 'Europe/Zurich' } },
    { label: 'Events with specific title', params: { meetingTitle: 'standup', startDateTime: '2026-03-18T00:00:00', endDateTime: '2026-03-25T00:00:00', timeZone: 'Europe/Zurich' } }
  ],
  ListCalendarView: [
    { label: 'Calendar view next 7 days', params: { userIdentifier: 'me', startDateTime: '2026-03-18T00:00:00', endDateTime: '2026-03-25T00:00:00', timeZone: 'Europe/Zurich' } },
    { label: 'Tomorrow expanded', params: { userIdentifier: 'me', startDateTime: '2026-03-19T00:00:00', endDateTime: '2026-03-20T00:00:00', timeZone: 'Europe/Zurich' } },
    { label: 'Next month overview', params: { userIdentifier: 'me', startDateTime: '2026-04-01T00:00:00', endDateTime: '2026-04-30T23:59:59', timeZone: 'Europe/Zurich', top: '50' } }
  ],
  CreateEvent: [
    { label: 'Quick 30min meeting', params: { subject: 'Quick sync', attendeeEmails: ['colleague@contoso.com'], startDateTime: '2026-03-19T10:00:00', endDateTime: '2026-03-19T10:30:00', timeZone: 'Europe/Zurich' } },
    { label: '1h Teams meeting with body', params: { subject: 'Project review', attendeeEmails: ['colleague@contoso.com'], startDateTime: '2026-03-20T14:00:00', endDateTime: '2026-03-20T15:00:00', timeZone: 'Europe/Zurich', isOnlineMeeting: 'true', bodyContent: 'Agenda: review progress and next steps' } },
    { label: 'All-day event', params: { subject: 'Team offsite', attendeeEmails: ['team@contoso.com'], startDateTime: '2026-03-21T00:00:00', endDateTime: '2026-03-22T00:00:00', timeZone: 'Europe/Zurich', location: 'Conference Room A' } }
  ],
  UpdateEvent: [
    { label: 'Change subject', params: { eventId: 'event-id-here', subject: 'Updated: Project review' } },
    { label: 'Reschedule to next week', params: { eventId: 'event-id-here', startDateTime: '2026-03-25T14:00:00', endDateTime: '2026-03-25T15:00:00', timeZone: 'Europe/Zurich' } }
  ],
  DeleteEventById: [
    { label: 'Delete by ID', params: { eventId: 'event-id-here' } }
  ],
  FindMeetingTimes: [
    { label: 'Find 1h slot', params: { userIdentifier: 'me', attendeeEmails: ['colleague@contoso.com'], meetingDuration: 'PT1H', timeZone: 'Europe/Zurich' } },
    { label: 'Find 30min slot this week', params: { userIdentifier: 'me', attendeeEmails: ['colleague@contoso.com'], meetingDuration: 'PT30M', startDateTime: '2026-03-18T08:00:00', endDateTime: '2026-03-21T18:00:00', timeZone: 'Europe/Zurich', maxCandidates: '5' } },
    { label: 'Find slot for 3 people', params: { userIdentifier: 'me', attendeeEmails: ['person1@contoso.com', 'person2@contoso.com'], meetingDuration: 'PT1H', timeZone: 'Europe/Zurich' } }
  ],
  AcceptEvent: [
    { label: 'Accept with comment', params: { eventId: 'event-id-here', comment: 'Looking forward to it!', sendResponse: 'true' } },
    { label: 'Accept silently', params: { eventId: 'event-id-here', sendResponse: 'false' } }
  ],
  TentativelyAcceptEvent: [
    { label: 'Tentative with note', params: { eventId: 'event-id-here', comment: 'Might have a conflict, will confirm', sendResponse: 'true' } }
  ],
  DeclineEvent: [
    { label: 'Decline with reason', params: { eventId: 'event-id-here', comment: 'Sorry, I have a conflict at that time', sendResponse: 'true' } }
  ],
  CancelEvent: [
    { label: 'Cancel with message', params: { eventId: 'event-id-here', comment: 'Meeting cancelled — will reschedule next week' } }
  ],
  ForwardEvent: [
    { label: 'Forward to colleague', params: { eventId: 'event-id-here', recipientEmails: ['colleague@contoso.com'], comment: 'FYI — thought you might want to join' } }
  ],
  GetRooms: [
    { label: 'List all rooms', params: {} }
  ],
  GetUserDateAndTimeZoneSettings: [
    { label: 'My timezone', params: { userIdentifier: 'me' } }
  ]
};
