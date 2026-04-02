#!/usr/bin/env node
// src/index.ts
// tva

import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { z } from 'zod';
import { TokenManager } from './auth.js';
import { CalendarClient } from './calendar.js';

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

    if (events.length === 0) {
      return {
        content: [{ type: 'text', text: 'No events found in the specified time range.' }],
      };
    }

    const lines = events.map((e) => {
      const time = e.isAllDay
        ? `All day`
        : `${e.start} → ${e.end}`;
      const flags = [
        e.isOnlineMeeting ? 'Teams' : '',
        e.isRecurring ? 'Recurring' : '',
        e.isPrivate ? 'Private' : '',
        e.showAs !== 'Busy' ? e.showAs : '',
      ]
        .filter(Boolean)
        .join(', ');
      return [
        `**${e.subject}**`,
        `  Time: ${time}`,
        e.organizer ? `  Organizer: ${e.organizer}` : '',
        e.location ? `  Location: ${e.location}` : '',
        flags ? `  Flags: ${flags}` : '',
        e.bodyPreview ? `  Preview: ${e.bodyPreview.substring(0, 120)}` : '',
      ]
        .filter(Boolean)
        .join('\n');
    });

    return {
      content: [
        {
          type: 'text',
          text: `Found ${events.length} event(s):\n\n${lines.join('\n\n')}`,
        },
      ],
    };
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
