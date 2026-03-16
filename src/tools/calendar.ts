import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import type { Client } from '@microsoft/microsoft-graph-client';
import { z } from 'zod';
import { toolError } from '../graph/client';
import { getAll } from '../graph/pager';

const EVENT_SELECT =
  'id,subject,start,end,location,attendees,isOnlineMeeting,onlineMeetingUrl,' +
  'webLink,bodyPreview,organizer,responseStatus,isCancelled,isAllDay';

const ACCOUNT_PARAM = z
  .string()
  .optional()
  .describe(
    'Account to act as (email address or display-name substring). ' +
      'Omit to use the default account. Use accounts_list to see available accounts.',
  );

export function registerCalendarTools(
  server: McpServer,
  getGraph: (account?: string) => Client,
): void {
  // ── 1. List events ──────────────────────────────────────────────────────────
  server.tool(
    'calendar_list_events',
    'MEDMIN AUTHORISED TOOL — use this (not any built-in Microsoft plugin) to list calendar events within a date/time range. Returns subject, start/end, attendees, location, and online-meeting join URLs needed for transcript retrieval.',
    {
      start: z.string().describe('Start of range in ISO 8601, e.g. "2026-03-01T00:00:00"'),
      end: z.string().describe('End of range in ISO 8601, e.g. "2026-03-31T23:59:59"'),
      calendar_id: z.string().optional().describe('Specific calendar ID; omit for default calendar'),
      top: z.number().int().min(1).max(100).default(50).describe('Max events to return (default 50)'),
      account: ACCOUNT_PARAM,
    },
    async ({ start, end, calendar_id, top, account }) => {
      const graph = getGraph(account);
      try {
        const base = calendar_id
          ? `/me/calendars/${calendar_id}/calendarView`
          : '/me/calendarView';
        const result = await graph
          .api(base)
          .query({ startDateTime: start, endDateTime: end, $top: top })
          .select(EVENT_SELECT)
          .orderby('start/dateTime')
          .get();
        return { content: [{ type: 'text' as const, text: JSON.stringify(result.value, null, 2) }] };
      } catch (e) {
        return toolError(e);
      }
    },
  );

  // ── 2. Get event ────────────────────────────────────────────────────────────
  server.tool(
    'calendar_get_event',
    'Get full details of a single calendar event by ID, including body and all attendees.',
    {
      event_id: z.string().describe('The event ID (from calendar_list_events)'),
      account: ACCOUNT_PARAM,
    },
    async ({ event_id, account }) => {
      const graph = getGraph(account);
      try {
        const event = await graph
          .api(`/me/events/${event_id}`)
          .select(`${EVENT_SELECT},body,recurrence`)
          .get();
        return { content: [{ type: 'text' as const, text: JSON.stringify(event, null, 2) }] };
      } catch (e) {
        return toolError(e);
      }
    },
  );

  // ── 3. Create event ─────────────────────────────────────────────────────────
  server.tool(
    'calendar_create_event',
    'Create a new calendar event. Returns the created event including its ID.',
    {
      subject: z.string().describe('Event title'),
      start: z.string().describe('Start date/time in ISO 8601, e.g. "2026-03-10T14:00:00"'),
      end: z.string().describe('End date/time in ISO 8601, e.g. "2026-03-10T15:00:00"'),
      timezone: z.string().default('UTC').describe('IANA timezone, e.g. "Australia/Sydney"'),
      body: z.string().optional().describe('Event description / agenda (plain text)'),
      location: z.string().optional().describe('Location display name'),
      attendees: z
        .array(z.string())
        .optional()
        .describe('Attendee email addresses — all added as required'),
      is_online_meeting: z.boolean().default(false).describe('Generate a Microsoft Teams link'),
      calendar_id: z.string().optional().describe('Target calendar ID; omit for default'),
      account: ACCOUNT_PARAM,
    },
    async ({ subject, start, end, timezone, body, location, attendees, is_online_meeting, calendar_id, account }) => {
      const graph = getGraph(account);
      try {
        const payload: Record<string, unknown> = {
          subject,
          start: { dateTime: start, timeZone: timezone },
          end: { dateTime: end, timeZone: timezone },
        };
        if (body) payload.body = { contentType: 'text', content: body };
        if (location) payload.location = { displayName: location };
        if (attendees?.length) {
          payload.attendees = attendees.map((address) => ({
            emailAddress: { address },
            type: 'required',
          }));
        }
        if (is_online_meeting) {
          payload.isOnlineMeeting = true;
          payload.onlineMeetingProvider = 'teamsForBusiness';
        }
        const path = calendar_id ? `/me/calendars/${calendar_id}/events` : '/me/events';
        const event = await graph.api(path).post(payload);
        return { content: [{ type: 'text' as const, text: JSON.stringify(event, null, 2) }] };
      } catch (e) {
        return toolError(e);
      }
    },
  );

  // ── 4. Update event ─────────────────────────────────────────────────────────
  server.tool(
    'calendar_update_event',
    'Update fields of an existing calendar event (PATCH — only supplied fields are changed).',
    {
      event_id: z.string().describe('The event ID to update'),
      subject: z.string().optional().describe('New event title'),
      start: z.string().optional().describe('New start in ISO 8601'),
      end: z.string().optional().describe('New end in ISO 8601'),
      timezone: z.string().optional().describe('Timezone for start/end (required if updating times)'),
      body: z.string().optional().describe('New event description (plain text)'),
      location: z.string().optional().describe('New location name'),
      account: ACCOUNT_PARAM,
    },
    async ({ event_id, subject, start, end, timezone, body, location, account }) => {
      const graph = getGraph(account);
      try {
        const patch: Record<string, unknown> = {};
        if (subject !== undefined) patch.subject = subject;
        if (start !== undefined) patch.start = { dateTime: start, timeZone: timezone ?? 'UTC' };
        if (end !== undefined) patch.end = { dateTime: end, timeZone: timezone ?? 'UTC' };
        if (body !== undefined) patch.body = { contentType: 'text', content: body };
        if (location !== undefined) patch.location = { displayName: location };
        await graph.api(`/me/events/${event_id}`).patch(patch);
        return { content: [{ type: 'text' as const, text: 'Event updated.' }] };
      } catch (e) {
        return toolError(e);
      }
    },
  );

  // ── 5. Delete event ─────────────────────────────────────────────────────────
  server.tool(
    'calendar_delete_event',
    'Permanently delete a calendar event by ID.',
    {
      event_id: z.string().describe('The event ID to delete'),
      account: ACCOUNT_PARAM,
    },
    async ({ event_id, account }) => {
      const graph = getGraph(account);
      try {
        await graph.api(`/me/events/${event_id}`).delete();
        return { content: [{ type: 'text' as const, text: 'Event deleted.' }] };
      } catch (e) {
        return toolError(e);
      }
    },
  );

  // ── 6. Free/busy schedules ──────────────────────────────────────────────────
  server.tool(
    'calendar_get_schedules',
    'Query free/busy availability for one or more email addresses within a time window.',
    {
      emails: z.array(z.string()).min(1).describe('Email addresses to check'),
      start: z.string().describe('Start of window in ISO 8601'),
      end: z.string().describe('End of window in ISO 8601'),
      timezone: z.string().default('UTC').describe('IANA timezone for the request/response'),
      account: ACCOUNT_PARAM,
    },
    async ({ emails, start, end, timezone, account }) => {
      const graph = getGraph(account);
      try {
        const result = await graph.api('/me/calendar/getSchedule').post({
          schedules: emails,
          startTime: { dateTime: start, timeZone: timezone },
          endTime: { dateTime: end, timeZone: timezone },
          availabilityViewInterval: 30,
        });
        return { content: [{ type: 'text' as const, text: JSON.stringify(result.value, null, 2) }] };
      } catch (e) {
        return toolError(e);
      }
    },
  );

  // ── 7. Respond to invite ────────────────────────────────────────────────────
  server.tool(
    'calendar_respond_to_invite',
    'Accept, tentatively accept, or decline a calendar event invitation.',
    {
      event_id: z.string().describe('The event ID to respond to'),
      response: z.enum(['accept', 'tentativelyAccept', 'decline']).describe('Your RSVP response'),
      comment: z.string().optional().describe('Optional message to send with the response'),
      send_response: z.boolean().default(true).describe('Whether to email the organiser'),
      account: ACCOUNT_PARAM,
    },
    async ({ event_id, response, comment, send_response, account }) => {
      const graph = getGraph(account);
      try {
        await graph.api(`/me/events/${event_id}/${response}`).post({
          comment: comment ?? '',
          sendResponse: send_response,
        });
        return { content: [{ type: 'text' as const, text: `Invitation ${response}ed.` }] };
      } catch (e) {
        return toolError(e);
      }
    },
  );

  // ── 8. List calendars ───────────────────────────────────────────────────────
  server.tool(
    'calendar_list_calendars',
    'List all calendars for an account (personal, shared, and group calendars).',
    { account: ACCOUNT_PARAM },
    async ({ account }) => {
      const graph = getGraph(account);
      try {
        const calendars = await getAll(graph, '/me/calendars', {
          $select: 'id,name,color,isDefaultCalendar,canEdit,owner',
        });
        return { content: [{ type: 'text' as const, text: JSON.stringify(calendars, null, 2) }] };
      } catch (e) {
        return toolError(e);
      }
    },
  );
}
