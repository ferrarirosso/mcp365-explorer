import * as React from 'react';
import {
  PrimaryButton,
  DefaultButton,
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
}

interface IEventItem {
  subject: string;
  start: string;
  end: string;
  location: string;
  isOnlineMeeting: boolean;
}

interface IShowcaseState {
  todayEvents: IEventItem[];
  weekEvents: IEventItem[];
  calendarViewEvents: IEventItem[];
  rooms: Array<{ displayName: string; emailAddress: string }>;
  timezone: string | undefined;
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
    if (outer && typeof outer === 'object' && typeof outer.response === 'string') {
      try { return JSON.parse(outer.response); } catch { return outer; }
    }
    return outer;
  } catch { return undefined; }
}

function extractEvents(result: IMcpCallResult): IEventItem[] {
  const parsed = extractJsonFromContent(result);
  if (!parsed) return [];
  let arr: unknown[];
  if (Array.isArray(parsed)) { arr = parsed; }
  else if (parsed && typeof parsed === 'object') {
    const obj = parsed as Record<string, unknown>;
    arr = Array.isArray(obj.value) ? obj.value : [obj];
  } else { return []; }

  return arr.map((e: unknown) => {
    const ev = e as Record<string, unknown>;
    const startObj = ev.start as Record<string, unknown> | undefined;
    const endObj = ev.end as Record<string, unknown> | undefined;
    const loc = ev.location as Record<string, unknown> | undefined;
    return {
      subject: String(ev.subject || ''),
      start: String(startObj?.dateTime || '').substring(0, 16),
      end: String(endObj?.dateTime || '').substring(0, 16),
      location: String(loc?.displayName || ''),
      isOnlineMeeting: !!ev.isOnlineMeeting
    };
  });
}

function todayRange(): { start: string; end: string } {
  const now = new Date();
  const start = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 0, 0, 0);
  const end = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59);
  return { start: start.toISOString(), end: end.toISOString() };
}

function weekRange(): { start: string; end: string } {
  const now = new Date();
  const start = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 0, 0, 0);
  const end = new Date(start.getTime() + 7 * 24 * 60 * 60 * 1000);
  return { start: start.toISOString(), end: end.toISOString() };
}

export const CalendarShowcase: React.FC<IShowcaseProps> = ({ client, theme }) => {
  const [state, setState] = React.useState<IShowcaseState>({
    todayEvents: [], weekEvents: [], calendarViewEvents: [], rooms: [], timezone: undefined, loading: undefined, error: undefined
  });

  const isLoading = !!state.loading;

  const handleToday = React.useCallback(async (): Promise<void> => {
    setState(prev => ({ ...prev, loading: 'ListEvents-today', error: undefined }));
    try {
      const range = todayRange();
      const result = await client.callTool('ListEvents', { startDateTime: range.start, endDateTime: range.end });
      setState(prev => ({ ...prev, loading: undefined, todayEvents: extractEvents(result) }));
    } catch (err) { setState(prev => ({ ...prev, loading: undefined, error: (err as Error).message })); }
  }, [client]);

  const handleWeek = React.useCallback(async (): Promise<void> => {
    setState(prev => ({ ...prev, loading: 'ListEvents-week', error: undefined }));
    try {
      const range = weekRange();
      const result = await client.callTool('ListEvents', { startDateTime: range.start, endDateTime: range.end });
      setState(prev => ({ ...prev, loading: undefined, weekEvents: extractEvents(result) }));
    } catch (err) { setState(prev => ({ ...prev, loading: undefined, error: (err as Error).message })); }
  }, [client]);

  const handleCalendarView = React.useCallback(async (): Promise<void> => {
    setState(prev => ({ ...prev, loading: 'ListCalendarView', error: undefined }));
    try {
      const result = await client.callTool('ListCalendarView', {});
      setState(prev => ({ ...prev, loading: undefined, calendarViewEvents: extractEvents(result) }));
    } catch (err) { setState(prev => ({ ...prev, loading: undefined, error: (err as Error).message })); }
  }, [client]);

  const handleRooms = React.useCallback(async (): Promise<void> => {
    setState(prev => ({ ...prev, loading: 'GetRooms', error: undefined }));
    try {
      const result = await client.callTool('GetRooms', {});
      const parsed = extractJsonFromContent(result);
      let rooms: Array<{ displayName: string; emailAddress: string }> = [];
      if (parsed && typeof parsed === 'object') {
        const obj = parsed as Record<string, unknown>;
        const arr = Array.isArray(obj.value) ? obj.value : Array.isArray(parsed) ? parsed as unknown[] : [];
        rooms = arr.map((r: unknown) => {
          const room = r as Record<string, unknown>;
          const addr = room.emailAddress as Record<string, unknown> | undefined;
          return { displayName: String(addr?.name || room.displayName || ''), emailAddress: String(addr?.address || room.emailAddress || '') };
        });
      }
      setState(prev => ({ ...prev, loading: undefined, rooms }));
    } catch (err) { setState(prev => ({ ...prev, loading: undefined, error: (err as Error).message })); }
  }, [client]);

  const handleTimezone = React.useCallback(async (): Promise<void> => {
    setState(prev => ({ ...prev, loading: 'GetUserDateAndTimeZoneSettings', error: undefined }));
    try {
      const result = await client.callTool('GetUserDateAndTimeZoneSettings', { userIdentifier: 'me' });
      const parsed = extractJsonFromContent(result) as Record<string, unknown> | undefined;
      setState(prev => ({ ...prev, loading: undefined, timezone: parsed?.timeZone as string || JSON.stringify(parsed) }));
    } catch (err) { setState(prev => ({ ...prev, loading: undefined, error: (err as Error).message })); }
  }, [client]);

  const cardStyle: React.CSSProperties = {
    border: `1px solid ${theme.palette.neutralLight}`, borderRadius: 8, padding: 12, backgroundColor: theme.palette.white, marginTop: 8
  };

  const renderEvents = (events: IEventItem[], label: string): React.ReactElement | null => {
    if (events.length === 0) return null;
    return (
      <div style={cardStyle}>
        <Text variant="small" styles={{ root: { fontWeight: 600, color: theme.palette.neutralSecondary, marginBottom: 8, display: 'block' } }}>
          {label} — {events.length} event{events.length !== 1 ? 's' : ''}
        </Text>
        <Stack tokens={{ childrenGap: 6 }}>
          {events.map((ev, i) => (
            <Stack key={i} horizontal verticalAlign="center" tokens={{ childrenGap: 8 }} styles={{ root: { padding: '4px 0', borderBottom: i < events.length - 1 ? `1px solid ${theme.palette.neutralLighter}` : 'none' } }}>
              <Icon iconName={ev.isOnlineMeeting ? 'TeamsLogo16' : 'Calendar'} styles={{ root: { color: theme.palette.themePrimary, fontSize: 14 } }} />
              <Stack>
                <Text styles={{ root: { fontWeight: 600, fontSize: 13 } }}>{ev.subject}</Text>
                <Text variant="small" styles={{ root: { color: theme.palette.neutralSecondary } }}>
                  {ev.start} — {ev.end}{ev.location ? ` · ${ev.location}` : ''}
                </Text>
              </Stack>
            </Stack>
          ))}
        </Stack>
      </div>
    );
  };

  return (
    <div style={{ marginTop: 12 }}>
      <Text variant="small" styles={{ root: { color: theme.palette.neutralTertiary, marginBottom: 12, display: 'block' } }}>
        Click the buttons below to query your calendar using MCP tools.
      </Text>

      {state.error && (
        <Text variant="small" styles={{ root: { color: theme.palette.red, marginBottom: 8, display: 'block' } }}>{state.error}</Text>
      )}

      <Stack horizontal tokens={{ childrenGap: 8 }} wrap styles={{ root: { marginBottom: 12 } }}>
        <PrimaryButton text={state.loading === 'ListEvents-today' ? 'Loading...' : 'My Events Today'} iconProps={{ iconName: 'CalendarDay' }} onClick={handleToday} disabled={isLoading} />
        <DefaultButton text={state.loading === 'ListEvents-week' ? 'Loading...' : 'My Events This Week'} iconProps={{ iconName: 'CalendarWeek' }} onClick={handleWeek} disabled={isLoading} />
        <DefaultButton text={state.loading === 'ListCalendarView' ? 'Loading...' : 'Calendar View (next 15 days)'} iconProps={{ iconName: 'CalendarMirrored' }} onClick={handleCalendarView} disabled={isLoading} />
        <DefaultButton text={state.loading === 'GetRooms' ? 'Loading...' : 'Meeting Rooms'} iconProps={{ iconName: 'Room' }} onClick={handleRooms} disabled={isLoading} />
        <DefaultButton text={state.loading === 'GetUserDateAndTimeZoneSettings' ? 'Loading...' : 'My Timezone'} iconProps={{ iconName: 'Globe' }} onClick={handleTimezone} disabled={isLoading} />
      </Stack>

      {isLoading && <Spinner size={SpinnerSize.small} label={`Calling ${state.loading}...`} styles={{ root: { marginBottom: 8 } }} />}

      {renderEvents(state.todayEvents, "Today's Events (ListEvents)")}
      {renderEvents(state.weekEvents, 'This Week (ListEvents)')}
      {renderEvents(state.calendarViewEvents, 'Calendar View — next 15 days (ListCalendarView) — recurring events expanded into individual instances')}

      {state.rooms.length > 0 && (
        <div style={cardStyle}>
          <Text variant="small" styles={{ root: { fontWeight: 600, color: theme.palette.neutralSecondary, marginBottom: 8, display: 'block' } }}>
            Meeting Rooms (GetRooms) — {state.rooms.length} room{state.rooms.length !== 1 ? 's' : ''}
          </Text>
          <Stack tokens={{ childrenGap: 4 }}>
            {state.rooms.map((room, i) => (
              <Stack key={i} horizontal tokens={{ childrenGap: 8 }}>
                <Icon iconName="Room" styles={{ root: { color: theme.palette.themePrimary } }} />
                <Text variant="small">{room.displayName}</Text>
                <Text variant="small" styles={{ root: { color: theme.palette.neutralTertiary } }}>{room.emailAddress}</Text>
              </Stack>
            ))}
          </Stack>
        </div>
      )}

      {state.timezone && (
        <div style={cardStyle}>
          <Text variant="small" styles={{ root: { fontWeight: 600, color: theme.palette.neutralSecondary } }}>
            My Timezone (GetUserDateAndTimeZoneSettings)
          </Text>
          <Text styles={{ root: { marginTop: 4 } }}>{state.timezone}</Text>
        </div>
      )}
    </div>
  );
};
