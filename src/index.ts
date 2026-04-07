#!/usr/bin/env node
// src/index.ts
// tva

import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { z } from 'zod';
import { TokenManager } from './auth.js';
import { CalendarClient } from './calendar.js';
import { MailClient } from './mail.js';
import type { OwaCreateEventPayload, OwaUpdateEventPayload, RsvpAction, OwaRsvpPayload, OwaRecipient, OwaUpdateMailPayload } from './types.js';

const tokenManager = new TokenManager();
const calendarClient = new CalendarClient(tokenManager);
const mailClient = new MailClient(tokenManager);

const server = new McpServer({
  name: 'owa-mcp',
  version: '0.4.1',
});

const recurrencePatternSchema = z.object({
  type: z.enum(['daily', 'weekly', 'absoluteMonthly', 'relativeMonthly', 'absoluteYearly', 'relativeYearly'])
    .describe('Recurrence pattern type'),
  interval: z.number().int().min(1).describe('Interval between occurrences (e.g. 1 = every week, 2 = every other week)'),
  daysOfWeek: z.array(z.string()).optional()
    .describe('Days for weekly/relative patterns: "Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"'),
  dayOfMonth: z.number().int().min(1).max(31).optional()
    .describe('Day of month for absoluteMonthly/absoluteYearly'),
  month: z.number().int().min(1).max(12).optional()
    .describe('Month (1-12) for yearly patterns'),
  index: z.enum(['first', 'second', 'third', 'fourth', 'last']).optional()
    .describe('Week index for relativeMonthly/relativeYearly'),
  firstDayOfWeek: z.string().optional()
    .describe('First day of the week (default "Sunday")'),
});

const recurrenceRangeSchema = z.object({
  type: z.enum(['endDate', 'numbered', 'noEnd'])
    .describe('"endDate" = recur until date, "numbered" = fixed count, "noEnd" = forever'),
  startDate: z.string().describe('Series start date in YYYY-MM-DD format'),
  endDate: z.string().optional().describe('End date (YYYY-MM-DD) — required when type is "endDate"'),
  numberOfOccurrences: z.number().int().min(1).optional()
    .describe('Number of occurrences — required when type is "numbered"'),
  recurrenceTimeZone: z.string().optional()
    .describe('Timezone for recurrence dates (e.g. "W. Europe Standard Time")'),
});

const recurrenceSchema = z.object({
  pattern: recurrencePatternSchema,
  range: recurrenceRangeSchema,
}).describe('Recurrence definition with pattern (how often) and range (how long)');

function toOwaRecurrence(rec: z.infer<typeof recurrenceSchema>): unknown {
  return {
    Pattern: {
      Type: rec.pattern.type.charAt(0).toUpperCase() + rec.pattern.type.slice(1),
      Interval: rec.pattern.interval,
      ...(rec.pattern.daysOfWeek && { DaysOfWeek: rec.pattern.daysOfWeek }),
      ...(rec.pattern.dayOfMonth !== undefined && { DayOfMonth: rec.pattern.dayOfMonth }),
      ...(rec.pattern.month !== undefined && { Month: rec.pattern.month }),
      ...(rec.pattern.index && { Index: rec.pattern.index.charAt(0).toUpperCase() + rec.pattern.index.slice(1) }),
      ...(rec.pattern.firstDayOfWeek && { FirstDayOfWeek: rec.pattern.firstDayOfWeek }),
    },
    Range: {
      Type: rec.range.type.charAt(0).toUpperCase() + rec.range.type.slice(1),
      StartDate: rec.range.startDate,
      ...(rec.range.endDate && { EndDate: rec.range.endDate }),
      ...(rec.range.numberOfOccurrences !== undefined && { NumberOfOccurrences: rec.range.numberOfOccurrences }),
      ...(rec.range.recurrenceTimeZone && { RecurrenceTimeZone: rec.range.recurrenceTimeZone }),
    },
  };
}

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
    hideAttendees: z.boolean().optional().default(false)
      .describe('Hide the attendee list so attendees cannot see who else was invited'),
    responseRequested: z.boolean().optional().default(true)
      .describe('Request attendees to send a response. Set false to not request RSVPs'),
    reminderMinutes: z.number().int().min(0).optional()
      .describe('Reminder in minutes before event start (e.g. 15). Omit to use Outlook default. Set to 0 to disable reminder'),
    recurrence: recurrenceSchema.optional()
      .describe('Make this a recurring event. Omit for a single event.'),
  },  async (params) => {
    const payload: OwaCreateEventPayload = {
      Subject: params.subject,
      Start: { DateTime: params.startDateTime, TimeZone: params.timezone },
      End: { DateTime: params.endDateTime, TimeZone: params.timezone },
      IsAllDay: params.isAllDay,
      ShowAs: params.showAs,
      Importance: params.importance,
      Sensitivity: params.isPrivate ? 'Private' : 'Normal',
      IsOnlineMeeting: params.isOnlineMeeting,
      HideAttendees: params.hideAttendees,
      ResponseRequested: params.responseRequested,
    };
    if (params.reminderMinutes !== undefined) {
      payload.IsReminderOn = params.reminderMinutes > 0;
      payload.ReminderMinutesBeforeStart = params.reminderMinutes;
    }
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
    if (params.recurrence) {
      payload.Recurrence = toOwaRecurrence(params.recurrence);
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
    hideAttendees: z.boolean().optional()
      .describe('Hide the attendee list so attendees cannot see who else was invited'),
    responseRequested: z.boolean().optional()
      .describe('Request attendees to send a response'),
    reminderMinutes: z.number().int().min(0).optional()
      .describe('Reminder in minutes before event start. Set to 0 to disable reminder'),
    recurrence: recurrenceSchema.optional()
      .describe('Change the recurrence pattern. Only applies to series master events.'),
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
    if (params.hideAttendees !== undefined) payload.HideAttendees = params.hideAttendees;
    if (params.responseRequested !== undefined) payload.ResponseRequested = params.responseRequested;
    if (params.reminderMinutes !== undefined) {
      payload.IsReminderOn = params.reminderMinutes > 0;
      payload.ReminderMinutesBeforeStart = params.reminderMinutes;
    }
    if (params.recurrence) {
      payload.Recurrence = toOwaRecurrence(params.recurrence);
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
    scope: z.enum(['single', 'thisAndFollowing', 'allInSeries']).optional().default('single')
      .describe('Scope of cancellation: "single" (default) cancels this occurrence only, "thisAndFollowing" cancels this and all future occurrences, "allInSeries" cancels the entire series'),
  },
  async ({ eventId, reason, scope }) => {
    await calendarClient.cancelEvent(eventId, reason, scope);
    return { content: [{ type: 'text', text: JSON.stringify({ cancelled: true, eventId, scope, reason: reason ?? null }, null, 2) }] };
  }
);

server.tool(
  'delete_calendar_event',
  'Remove an event from your calendar without sending any notification. Use this to remove events you did not organize, or to silently delete your own events.',
  {
    eventId: z.string().describe('Event ID from get_calendar_events'),
    scope: z.enum(['single', 'thisAndFollowing', 'allInSeries']).optional().default('single')
      .describe('Scope of deletion: "single" (default) deletes this occurrence only, "thisAndFollowing" deletes this and all future occurrences, "allInSeries" deletes the entire series'),
  },
  async ({ eventId, scope }) => {
    await calendarClient.deleteEvent(eventId, scope);
    return { content: [{ type: 'text', text: JSON.stringify({ deleted: true, eventId, scope }, null, 2) }] };
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
    await calendarClient.respondToEvent(params.eventId, actionMap[params.response], payload, params.proposedTimezone);
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

server.tool(
  'get_series_master',
  'Inspect the master event of a recurring series. Returns recurrence pattern, cancelled occurrences, and full event details. Accepts any event ID from the series (occurrence, exception, or master).',
  {
    eventId: z.string().describe('Any event ID from the series — occurrence, exception, or series master. Resolved automatically.'),
    timezone: z.string().optional().default('UTC')
      .describe('IANA timezone name for event times, e.g. Europe/Berlin'),
  },
  async ({ eventId, timezone }) => {
    const master = await calendarClient.getSeriesMaster(eventId, timezone);
    return { content: [{ type: 'text', text: JSON.stringify(master, null, 2) }] };
  }
);

server.tool(
  'list_series_instances',
  'List all occurrences of a recurring series within a date range. Accepts any event ID from the series (resolved to master automatically).',
  {
    eventId: z.string().describe('Any event ID from the series — occurrence, exception, or series master. Resolved automatically.'),
    startDateTime: z.string().describe('Start of time range in ISO 8601 format, e.g. 2026-04-07T00:00:00Z'),
    endDateTime: z.string().describe('End of time range in ISO 8601 format, e.g. 2026-07-07T00:00:00Z'),
    timezone: z.string().optional().default('UTC')
      .describe('IANA timezone name for event times, e.g. Europe/Berlin'),
  },
  async ({ eventId, startDateTime, endDateTime, timezone }) => {
    const instances = await calendarClient.listSeriesInstances(eventId, startDateTime, endDateTime, timezone);
    return { content: [{ type: 'text', text: JSON.stringify(instances, null, 2) }] };
  }
);

server.tool(
  'list_mail_folders',
  'List all mail folders in the mailbox, or child folders of a specific folder.',
  {
    parentFolderId: z.string().optional()
      .describe('List children of this folder. If omitted, lists top-level folders.'),
  },
  async ({ parentFolderId }) => {
    const folders = await mailClient.listFolders(parentFolderId);
    return { content: [{ type: 'text', text: JSON.stringify(folders, null, 2) }] };
  }
);

server.tool(
  'get_emails',
  'Get emails from a specific mailbox folder with optional filtering.',
  {
    folderId: z.string().optional().default('Inbox')
      .describe('Folder ID or well-known name (Inbox, Drafts, SentItems, DeletedItems). Default: Inbox'),
    filter: z.enum(['all', 'unread', 'flagged', 'today', 'this_week']).optional().default('all')
      .describe('Filter preset'),
    limit: z.number().int().min(1).max(500).optional().default(20)
      .describe('Maximum number of emails to return (default 20, max 500)'),
    pageToken: z.string().optional()
      .describe('Pagination token from previous response'),
  },
  async ({ folderId, filter, limit, pageToken }) => {
    const result = await mailClient.getMessages(folderId, { filter, limit, pageToken });
    return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
  }
);

server.tool(
  'search_emails',
  'Search emails using full-text query OR structured filters (mutually exclusive). Use query for natural search, or structured filters for precise field matching.',
  {
    query: z.string().optional()
      .describe('Full-text search query (uses Exchange search index). Cannot be combined with structured filters.'),
    from: z.string().optional()
      .describe('Filter by sender email address'),
    subject: z.string().optional()
      .describe('Filter by subject (contains match)'),
    receivedAfter: z.string().optional()
      .describe('ISO 8601 datetime — only messages received after this time'),
    receivedBefore: z.string().optional()
      .describe('ISO 8601 datetime — only messages received before this time'),
    folderId: z.string().optional()
      .describe('Scope search to a specific folder'),
    limit: z.number().int().min(1).max(500).optional().default(20)
      .describe('Maximum results (default 20, max 500)'),
  },
  async (params) => {
    const result = await mailClient.searchMessages(params);
    return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
  }
);

server.tool(
  'get_email',
  'Read a single email with full body content and attachment metadata.',
  {
    messageId: z.string().describe('Message ID from get_emails or search_emails'),
    format: z.enum(['text', 'html']).optional().default('text')
      .describe('Body format: "text" (default) or "html"'),
  },
  async ({ messageId, format }) => {
    const message = await mailClient.getMessage(messageId, format);
    return { content: [{ type: 'text', text: JSON.stringify(message, null, 2) }] };
  }
);

server.tool(
  'get_attachment',
  'Download an email attachment to disk. Returns the file path.',
  {
    messageId: z.string().describe('Message ID'),
    attachmentId: z.string().describe('Attachment ID from get_email response'),
  },
  async ({ messageId, attachmentId }) => {
    const result = await mailClient.getAttachment(messageId, attachmentId);
    return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
  }
);

const recipientSchema = z.object({
  email: z.string().describe('Email address'),
  name: z.string().optional().describe('Display name'),
});

function toOwaRecipients(arr: { email: string; name?: string }[]): OwaRecipient[] {
  return arr.map(r => ({ EmailAddress: { Address: r.email, Name: r.name } }));
}

server.tool(
  'send_email',
  'Compose and send a new email in one step. For more control (edit before sending), use create_draft + update_draft + send_draft instead.',
  {
    to: z.array(recipientSchema).min(1).describe('Recipients'),
    subject: z.string().describe('Email subject'),
    body: z.string().describe('Email body content'),
    bodyType: z.enum(['text', 'html']).optional().default('text')
      .describe('Body format: "text" (default) or "html"'),
    cc: z.array(recipientSchema).optional().describe('CC recipients'),
    bcc: z.array(recipientSchema).optional().describe('BCC recipients'),
    importance: z.enum(['Low', 'Normal', 'High']).optional().default('Normal'),
    saveToSentItems: z.boolean().optional().default(true)
      .describe('Save a copy in Sent Items (default true)'),
  },
  async (params) => {
    await mailClient.sendMail({
      Message: {
        Subject: params.subject,
        Body: { ContentType: params.bodyType === 'html' ? 'HTML' : 'Text', Content: params.body },
        ToRecipients: toOwaRecipients(params.to),
        CcRecipients: params.cc ? toOwaRecipients(params.cc) : undefined,
        BccRecipients: params.bcc ? toOwaRecipients(params.bcc) : undefined,
        Importance: params.importance,
      },
      SaveToSentItems: params.saveToSentItems,
    });
    return { content: [{ type: 'text', text: JSON.stringify({ sent: true }, null, 2) }] };
  }
);

server.tool(
  'create_draft',
  'Create a new email draft (saved to Drafts folder). Use update_draft to modify, then send_draft to send.',
  {
    to: z.array(recipientSchema).min(1).describe('Recipients'),
    subject: z.string().describe('Email subject'),
    body: z.string().describe('Email body content'),
    bodyType: z.enum(['text', 'html']).optional().default('text')
      .describe('Body format: "text" (default) or "html"'),
    cc: z.array(recipientSchema).optional().describe('CC recipients'),
    bcc: z.array(recipientSchema).optional().describe('BCC recipients'),
    importance: z.enum(['Low', 'Normal', 'High']).optional().default('Normal'),
  },
  async (params) => {
    const draft = await mailClient.createDraft({
      Subject: params.subject,
      Body: { ContentType: params.bodyType === 'html' ? 'HTML' : 'Text', Content: params.body },
      ToRecipients: toOwaRecipients(params.to),
      CcRecipients: params.cc ? toOwaRecipients(params.cc) : undefined,
      BccRecipients: params.bcc ? toOwaRecipients(params.bcc) : undefined,
      Importance: params.importance,
    });
    return { content: [{ type: 'text', text: JSON.stringify(draft, null, 2) }] };
  }
);

server.tool(
  'create_reply_draft',
  'Create a draft reply to the sender of a message. Returns the draft with pre-filled recipients, quoted body, and "RE:" subject. Use update_draft to modify, then send_draft to send.',
  {
    messageId: z.string().describe('Message ID from get_emails or search_emails'),
  },
  async ({ messageId }) => {
    const draft = await mailClient.createReplyDraft(messageId);
    return { content: [{ type: 'text', text: JSON.stringify(draft, null, 2) }] };
  }
);

server.tool(
  'create_reply_all_draft',
  'Create a draft reply-all to all recipients of a message. Returns the draft with pre-filled recipients, quoted body, and "RE:" subject. Use update_draft to modify, then send_draft to send.',
  {
    messageId: z.string().describe('Message ID from get_emails or search_emails'),
  },
  async ({ messageId }) => {
    const draft = await mailClient.createReplyAllDraft(messageId);
    return { content: [{ type: 'text', text: JSON.stringify(draft, null, 2) }] };
  }
);

server.tool(
  'create_forward_draft',
  'Create a draft forward of a message. Returns the draft with quoted body and "FW:" subject but no To recipients. Use update_draft to set recipients, then send_draft to send.',
  {
    messageId: z.string().describe('Message ID from get_emails or search_emails'),
  },
  async ({ messageId }) => {
    const draft = await mailClient.createForwardDraft(messageId);
    return { content: [{ type: 'text', text: JSON.stringify(draft, null, 2) }] };
  }
);

server.tool(
  'update_draft',
  'Modify a draft message before sending. Can change subject, body, recipients, and importance. Use with create_draft, create_reply_draft, create_reply_all_draft, or create_forward_draft.',
  {
    messageId: z.string().describe('Draft message ID'),
    subject: z.string().optional().describe('New subject'),
    body: z.string().optional().describe('New body content'),
    bodyType: z.enum(['text', 'html']).optional().describe('Body format'),
    toRecipients: z.array(recipientSchema).optional().describe('Replace all To recipients'),
    ccRecipients: z.array(recipientSchema).optional().describe('Replace all CC recipients'),
    bccRecipients: z.array(recipientSchema).optional().describe('Replace all BCC recipients'),
    importance: z.enum(['Low', 'Normal', 'High']).optional(),
  },
  async (params) => {
    const payload: OwaUpdateMailPayload = {};
    if (params.subject !== undefined) payload.Subject = params.subject;
    if (params.body !== undefined) {
      payload.Body = {
        ContentType: params.bodyType === 'html' ? 'HTML' : 'Text',
        Content: params.body,
      };
    }
    if (params.toRecipients !== undefined) payload.ToRecipients = toOwaRecipients(params.toRecipients);
    if (params.ccRecipients !== undefined) payload.CcRecipients = toOwaRecipients(params.ccRecipients);
    if (params.bccRecipients !== undefined) payload.BccRecipients = toOwaRecipients(params.bccRecipients);
    if (params.importance !== undefined) payload.Importance = params.importance;
    const updated = await mailClient.updateMessage(params.messageId, payload);
    return { content: [{ type: 'text', text: JSON.stringify(updated, null, 2) }] };
  }
);

server.tool(
  'send_draft',
  'Send a draft message. The draft is moved from Drafts to Sent Items.',
  {
    messageId: z.string().describe('Draft message ID from create_draft, create_reply_draft, create_reply_all_draft, or create_forward_draft'),
  },
  async ({ messageId }) => {
    await mailClient.sendDraft(messageId);
    return { content: [{ type: 'text', text: JSON.stringify({ sent: true, messageId }, null, 2) }] };
  }
);

server.tool(
  'move_email',
  'Move a message to a different folder. Returns the moved message (with updated ID).',
  {
    messageId: z.string().describe('Message ID'),
    destinationId: z.string().describe('Destination folder ID or well-known name (Inbox, Drafts, SentItems, DeletedItems, Archive)'),
  },
  async ({ messageId, destinationId }) => {
    const moved = await mailClient.moveMessage(messageId, destinationId);
    return { content: [{ type: 'text', text: JSON.stringify(moved, null, 2) }] };
  }
);

server.tool(
  'delete_email',
  'Delete a message (moves to Deleted Items).',
  {
    messageId: z.string().describe('Message ID'),
  },
  async ({ messageId }) => {
    await mailClient.deleteMessage(messageId);
    return { content: [{ type: 'text', text: JSON.stringify({ deleted: true, messageId }, null, 2) }] };
  }
);

server.tool(
  'update_email',
  'Update email properties: mark as read/unread, flag/unflag.',
  {
    messageId: z.string().describe('Message ID'),
    isRead: z.boolean().optional().describe('Set read (true) or unread (false)'),
    flagStatus: z.enum(['NotFlagged', 'Flagged', 'Complete']).optional()
      .describe('Flag status'),
  },
  async (params) => {
    const payload: OwaUpdateMailPayload = {};
    if (params.isRead !== undefined) payload.IsRead = params.isRead;
    if (params.flagStatus !== undefined) payload.Flag = { FlagStatus: params.flagStatus };
    const updated = await mailClient.updateMessage(params.messageId, payload);
    return { content: [{ type: 'text', text: JSON.stringify(updated, null, 2) }] };
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
