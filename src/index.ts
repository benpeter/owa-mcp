#!/usr/bin/env node
// src/index.ts
// tva

import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { z } from 'zod';
import { TokenManager } from './auth.js';
import { CalendarClient } from './calendar.js';
import type { OwaCreateEventPayload, OwaUpdateEventPayload, RsvpAction, OwaRsvpPayload } from './types.js';

const tokenManager = new TokenManager();
const calendarClient = new CalendarClient(tokenManager);

const server = new McpServer({
  name: 'owa-mcp',
  version: '0.1.0',
});

server.tool(
  'get_calendar_events',
  'Retrieve calendar events from Microsoft Outlook. Returns events between startDateTime and endDateTime.',
  {
    startDateTime: z
      .string()
      .describe('Start of time range in ISO 8601 format, e.g. 2026-04-07T00:00:00Z'),
    endDateTime: z
      .string()
      .describe('End of time range in ISO 8601 format, e.g. 2026-04-14T00:00:00Z'),
    maxResults: z
      .number()
      .int()
      .min(1)
      .max(100)
      .optional()
      .default(50)
      .describe('Maximum number of events to return (default 50, max 100)'),
    timezone: z
      .string()
      .optional()
      .default('UTC')
      .describe('IANA timezone name for event times, e.g. Europe/Berlin'),
  },
  async ({ startDateTime, endDateTime, maxResults, timezone }) => {
    const events = await calendarClient.getCalendarEvents(startDateTime, endDateTime, {
      maxResults,
      timezone,
    });

    return {
      content: [{ type: 'text', text: JSON.stringify(events, null, 2) }],
    };
  }
);

server.tool(
  'create_calendar_event',
  'Create a new calendar event in Microsoft Outlook. Adding attendees automatically sends meeting invitations.',
  {
    subject: z.string().describe('Event title'),
    startDateTime: z.string().describe('Start time as local datetime WITHOUT offset, e.g. 2026-04-07T09:00:00'),
    endDateTime: z.string().describe('End time as local datetime WITHOUT offset, e.g. 2026-04-07T10:00:00'),
    timezone: z.string().optional().default('W. Europe Standard Time')
      .describe('Windows timezone name for start/end times, e.g. "W. Europe Standard Time", "Pacific Standard Time", "UTC"'),
    body: z.string().optional().describe('Event body/description (plain text)'),
    location: z.string().optional().describe('Location name'),
    attendees: z.array(z.object({
      email: z.string().describe('Attendee email address'),
      name: z.string().optional().describe('Attendee display name'),
      type: z.enum(['Required', 'Optional', 'Resource']).optional().default('Required'),
    })).optional().describe('List of attendees. Adding attendees sends meeting invitations automatically.'),
    isAllDay: z.boolean().optional().default(false).describe('All-day event. If true, start/end must be midnight-to-midnight.'),
    showAs: z.enum(['Free', 'Tentative', 'Busy', 'Oof', 'WorkingElsewhere']).optional().default('Busy'),
    importance: z.enum(['Low', 'Normal', 'High']).optional().default('Normal'),
    isPrivate: z.boolean().optional().default(false),
    isOnlineMeeting: z.boolean().optional().default(false).describe('Create as Teams meeting'),
  },
  async (params) => {
    const payload: OwaCreateEventPayload = {
      Subject: params.subject,
      Start: { DateTime: params.startDateTime, TimeZone: params.timezone },
      End: { DateTime: params.endDateTime, TimeZone: params.timezone },
      IsAllDay: params.isAllDay,
      ShowAs: params.showAs,
      Importance: params.importance,
      Sensitivity: params.isPrivate ? 'Private' : 'Normal',
      IsOnlineMeeting: params.isOnlineMeeting,
    };
    if (params.body) {
      payload.Body = { ContentType: 'Text', Content: params.body };
    }
    if (params.location) {
      payload.Location = { DisplayName: params.location };
    }
    if (params.attendees) {
      payload.Attendees = params.attendees.map(a => ({
        EmailAddress: { Address: a.email, Name: a.name ?? a.email },
        Type: a.type ?? 'Required',
      }));
    }
    const event = await calendarClient.createEvent(payload, params.timezone);
    return { content: [{ type: 'text', text: JSON.stringify(event, null, 2) }] };
  }
);

server.tool(
  'update_calendar_event',
  'Update an existing calendar event. Only include fields you want to change. If you are the organizer, updates are sent to attendees automatically.',
  {
    eventId: z.string().describe('Event ID from get_calendar_events'),
    subject: z.string().optional(),
    startDateTime: z.string().optional().describe('Local datetime WITHOUT offset'),
    endDateTime: z.string().optional().describe('Local datetime WITHOUT offset'),
    timezone: z.string().optional().default('W. Europe Standard Time')
      .describe('Windows timezone name for start/end times'),
    body: z.string().optional().describe('New event body (plain text). WARNING: for online meetings, this overwrites the Teams join link.'),
    location: z.string().optional(),
    showAs: z.enum(['Free', 'Tentative', 'Busy', 'Oof', 'WorkingElsewhere']).optional(),
    isPrivate: z.boolean().optional(),
  },
  async (params) => {
    const payload: OwaUpdateEventPayload = {};
    if (params.subject !== undefined) payload.Subject = params.subject;
    if (params.startDateTime !== undefined) {
      payload.Start = { DateTime: params.startDateTime, TimeZone: params.timezone! };
    }
    if (params.endDateTime !== undefined) {
      payload.End = { DateTime: params.endDateTime, TimeZone: params.timezone! };
    }
    if (params.body !== undefined) {
      payload.Body = { ContentType: 'Text', Content: params.body };
    }
    if (params.location !== undefined) {
      payload.Location = { DisplayName: params.location };
    }
    if (params.showAs !== undefined) payload.ShowAs = params.showAs;
    if (params.isPrivate !== undefined) {
      payload.Sensitivity = params.isPrivate ? 'Private' : 'Normal';
    }
    const event = await calendarClient.updateEvent(params.eventId, payload, params.timezone);
    return { content: [{ type: 'text', text: JSON.stringify(event, null, 2) }] };
  }
);

server.tool(
  'cancel_calendar_event',
  'Cancel a meeting you organized. Sends a cancellation notice with your reason to all attendees. Only works if you are the organizer.',
  {
    eventId: z.string().describe('Event ID from get_calendar_events'),
    reason: z.string().optional().describe('Cancellation reason sent to attendees'),
  },
  async ({ eventId, reason }) => {
    await calendarClient.cancelEvent(eventId, reason);
    return { content: [{ type: 'text', text: JSON.stringify({ cancelled: true, eventId, reason: reason ?? null }, null, 2) }] };
  }
);

server.tool(
  'delete_calendar_event',
  'Remove an event from your calendar without sending any notification. Use this to remove events you did not organize, or to silently delete your own events.',
  {
    eventId: z.string().describe('Event ID from get_calendar_events'),
  },
  async ({ eventId }) => {
    await calendarClient.deleteEvent(eventId);
    return { content: [{ type: 'text', text: JSON.stringify({ deleted: true, eventId }, null, 2) }] };
  }
);

server.tool(
  'respond_to_calendar_event',
  'RSVP to a meeting invitation: accept, tentatively accept, or decline. Optionally include a comment and/or propose an alternative time.',
  {
    eventId: z.string().describe('Event ID from get_calendar_events'),
    response: z.enum(['accept', 'tentative', 'decline']).describe('Your response'),
    comment: z.string().optional().describe('Message sent to the organizer with your response'),
    sendResponse: z.boolean().optional().default(true)
      .describe('Whether to notify the organizer. Set false to RSVP silently.'),
    proposedStartDateTime: z.string().optional()
      .describe('Propose alternative start time (only for tentative/decline). Local datetime WITHOUT offset.'),
    proposedEndDateTime: z.string().optional()
      .describe('Propose alternative end time. Local datetime WITHOUT offset.'),
    proposedTimezone: z.string().optional().default('W. Europe Standard Time')
      .describe('Timezone for proposed times'),
  },
  async (params) => {
    const actionMap: Record<string, RsvpAction> = {
      accept: 'accept',
      tentative: 'tentativelyaccept',
      decline: 'decline',
    };
    const payload: OwaRsvpPayload = {
      SendResponse: params.sendResponse,
    };
    if (params.comment) payload.Comment = params.comment;
    if (params.proposedStartDateTime && params.proposedEndDateTime) {
      payload.ProposedNewTime = {
        Start: { DateTime: params.proposedStartDateTime, TimeZone: params.proposedTimezone! },
        End: { DateTime: params.proposedEndDateTime, TimeZone: params.proposedTimezone! },
      };
    }
    await calendarClient.respondToEvent(params.eventId, actionMap[params.response], payload);
    return {
      content: [{ type: 'text', text: JSON.stringify({ responded: true, eventId: params.eventId, response: params.response, comment: params.comment ?? null }, null, 2) }],
    };
  }
);

server.tool(
  'follow_calendar_event',
  'Follow a calendar event without formally RSVPing. The event appears on your calendar with ShowAs=Free. The organizer is NOT notified. Emulates New Outlook\'s "Follow this event" feature.',
  {
    eventId: z.string().describe('Event ID from get_calendar_events'),
    comment: z.string().optional().describe('Optional message included in the follow notification to the organizer'),
    timezone: z.string().optional().default('W. Europe Standard Time')
      .describe('Timezone for returned event times'),
  },
  async ({ eventId, comment, timezone }) => {
    const event = await calendarClient.followEvent(eventId, comment, timezone);
    return { content: [{ type: 'text', text: JSON.stringify(event, null, 2) }] };
  }
);

async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  // stderr so it doesn't interfere with MCP stdio protocol
  process.stderr.write('owa-mcp server running on stdio\n');
}

main().catch((err) => {
  process.stderr.write(`Fatal: ${err}\n`);
  process.exit(1);
});
