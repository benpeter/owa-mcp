# Calendar Write Tools — Create, Update, Cancel, RSVP, Follow

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add calendar write operations to the owa-mcp server: create events, update events, cancel/delete events, RSVP to invitations, and "follow" events without formally RSVPing.

**Builds on:** `2026-04-02-owa-mcp-foundation.md` — the existing MCP server with `get_calendar_events` and the CalendarClient/TokenManager architecture.

**API base:** `https://outlook.office.com/api/v2.0` (PascalCase properties). Token already carries `Calendars.ReadWrite` scope.

---

## Context

The foundation plan delivered a read-only `get_calendar_events` tool. Users need write operations to manage their calendar from Claude Code: creating meetings, rescheduling, cancelling with a reason, RSVPing to invitations, and "following" events they want to track without blocking time.

The OWA REST API v2.0 supports all of these except "follow", which is a New Outlook client-only feature with no public API. We emulate it via `tentativelyAccept(sendResponse=false)` + `PATCH showAs=Free`.

---

## New MCP Tools

| Tool | API Operation | Description |
|------|--------------|-------------|
| `create_calendar_event` | `POST /me/events` | Create a new event, optionally with attendees |
| `update_calendar_event` | `PATCH /me/events/{id}` | Update any fields on an existing event |
| `cancel_calendar_event` | `POST /me/events/{id}/cancel` | Organizer cancels with optional reason (sends message to attendees) |
| `delete_calendar_event` | `DELETE /me/events/{id}` | Remove event from calendar (no message sent) |
| `respond_to_calendar_event` | `POST /me/events/{id}/{action}` | RSVP: accept, tentativelyaccept, or decline with optional comment |
| `follow_calendar_event` | tentativelyAccept + PATCH | Track event on calendar without RSVPing or blocking time |

---

## Files to Modify/Create

| File | Change |
|------|--------|
| `src/types.ts` | Add write-related types: `CreateEventPayload`, `UpdateEventPayload`, `OwaEventResponse` |
| `src/calendar.ts` | Add methods: `createEvent`, `updateEvent`, `cancelEvent`, `deleteEvent`, `respondToEvent`, `followEvent` |
| `src/index.ts` | Register 6 new tools with Zod schemas |
| `tests/calendar-write.test.ts` | Integration tests for write operations |
| `CLAUDE.md` | Update Available Tools section |
| `README.md` | Update Available Tools table and Roadmap |

---

## Task 1: Add Write Types to `src/types.ts`

**Files:** Modify `src/types.ts`

- [ ] **Step 1: Add types for event creation and update**

Append to `src/types.ts`:

```typescript
// Request payload for POST /me/events (PascalCase for OWA API)
export interface OwaCreateEventPayload {
  Subject: string;
  Start: { DateTime: string; TimeZone: string };
  End: { DateTime: string; TimeZone: string };
  Body?: { ContentType: 'HTML' | 'Text'; Content: string };
  Location?: { DisplayName: string };
  Attendees?: OwaAttendee[];
  IsAllDay?: boolean;
  ShowAs?: string;
  Importance?: 'Low' | 'Normal' | 'High';
  Sensitivity?: 'Normal' | 'Personal' | 'Private' | 'Confidential';
  IsReminderOn?: boolean;
  ReminderMinutesBeforeStart?: number;
  IsOnlineMeeting?: boolean;
  Recurrence?: unknown;
}

export interface OwaAttendee {
  EmailAddress: { Address: string; Name?: string };
  Type: 'Required' | 'Optional' | 'Resource';
}

// Request payload for PATCH /me/events/{id}
// Same as create but everything optional
export type OwaUpdateEventPayload = Partial<OwaCreateEventPayload>;

// Response from POST/PATCH — full event object from OWA
// Reuse OwaCalendarEvent (it's the same shape the API returns)

// RSVP action types
export type RsvpAction = 'accept' | 'tentativelyaccept' | 'decline';

export interface OwaRsvpPayload {
  Comment?: string;
  SendResponse?: boolean;
  ProposedNewTime?: {
    Start: { DateTime: string; TimeZone: string };
    End: { DateTime: string; TimeZone: string };
  };
}
```

- [ ] **Step 2: Verify types compile**

```bash
npx tsc --noEmit
```

- [ ] **Step 3: Commit**

```bash
git add src/types.ts
git commit -m "feat: add write operation types for calendar events"
```

---

## Task 2: Add Write Methods to `src/calendar.ts`

**Files:** Modify `src/calendar.ts`

- [ ] **Step 1: Add helper method for authenticated API calls**

Add a private helper to CalendarClient that handles token, headers, and error checking. This replaces the inline fetch in `getCalendarEvents` and is reused by all new methods.

```typescript
private async request(
  method: string,
  path: string,
  options: { body?: unknown; timezone?: string } = {}
): Promise<Response> {
  const token = await this.tokens.getToken();
  const headers: Record<string, string> = {
    Authorization: `Bearer ${token.value}`,
    Accept: 'application/json',
  };
  if (options.timezone) {
    headers['Prefer'] = `outlook.timezone="${options.timezone}"`;
  }
  if (options.body !== undefined) {
    headers['Content-Type'] = 'application/json';
  }

  const url = path.startsWith('http') ? path : `${OWA_BASE}${path}`;
  const res = await fetch(url, {
    method,
    headers,
    body: options.body !== undefined ? JSON.stringify(options.body) : undefined,
  });

  if (!res.ok) {
    const text = await res.text();
    throw new Error(`OWA API error ${res.status} ${method} ${path}: ${text}`);
  }
  return res;
}
```

- [ ] **Step 2: Refactor `getCalendarEvents` to use the helper**

Replace the inline fetch loop in `getCalendarEvents` to use `this.request()` for the initial call and pagination. The logic stays the same — just use the helper for the HTTP calls.

- [ ] **Step 3: Add `createEvent` method**

```typescript
async createEvent(
  payload: OwaCreateEventPayload,
  timezone?: string
): Promise<CalendarEvent> {
  const res = await this.request('POST', '/me/events', {
    body: payload,
    timezone,
  });
  const raw = (await res.json()) as OwaCalendarEvent;
  return this.normalise(raw);
}
```

- [ ] **Step 4: Add `updateEvent` method**

```typescript
async updateEvent(
  eventId: string,
  payload: OwaUpdateEventPayload,
  timezone?: string
): Promise<CalendarEvent> {
  const res = await this.request('PATCH', `/me/events/${eventId}`, {
    body: payload,
    timezone,
  });
  const raw = (await res.json()) as OwaCalendarEvent;
  return this.normalise(raw);
}
```

- [ ] **Step 5: Add `cancelEvent` method**

```typescript
async cancelEvent(eventId: string, comment?: string): Promise<void> {
  await this.request('POST', `/me/events/${eventId}/cancel`, {
    body: comment ? { Comment: comment } : {},
  });
}
```

- [ ] **Step 6: Add `deleteEvent` method**

```typescript
async deleteEvent(eventId: string): Promise<void> {
  await this.request('DELETE', `/me/events/${eventId}`);
}
```

- [ ] **Step 7: Add `respondToEvent` method**

```typescript
async respondToEvent(
  eventId: string,
  action: RsvpAction,
  payload: OwaRsvpPayload = {}
): Promise<void> {
  await this.request('POST', `/me/events/${eventId}/${action}`, {
    body: payload,
  });
}
```

- [ ] **Step 8: Add `followEvent` method**

Emulates New Outlook's "Follow" by: (1) tentatively accepting without notifying the organizer, (2) setting ShowAs to Free.

```typescript
async followEvent(eventId: string, timezone?: string): Promise<CalendarEvent> {
  await this.request('POST', `/me/events/${eventId}/tentativelyaccept`, {
    body: { SendResponse: false },
  });
  const res = await this.request('PATCH', `/me/events/${eventId}`, {
    body: { ShowAs: 'Free' },
    timezone,
  });
  const raw = (await res.json()) as OwaCalendarEvent;
  return this.normalise(raw);
}
```

- [ ] **Step 9: Verify it compiles**

```bash
npx tsc --noEmit
```

- [ ] **Step 10: Commit**

```bash
git add src/calendar.ts
git commit -m "feat: CalendarClient write methods — create, update, cancel, delete, RSVP, follow"
```

---

## Task 3: Register New Tools in `src/index.ts`

**Files:** Modify `src/index.ts`

- [ ] **Step 1: Add `create_calendar_event` tool**

```typescript
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
```

- [ ] **Step 2: Add `update_calendar_event` tool**

```typescript
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
```

- [ ] **Step 3: Add `cancel_calendar_event` tool**

```typescript
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
```

- [ ] **Step 4: Add `delete_calendar_event` tool**

```typescript
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
```

- [ ] **Step 5: Add `respond_to_calendar_event` tool**

```typescript
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
```

- [ ] **Step 6: Add `follow_calendar_event` tool**

```typescript
server.tool(
  'follow_calendar_event',
  'Follow a calendar event without formally RSVPing. The event appears on your calendar with ShowAs=Free. The organizer is NOT notified. Emulates New Outlook\'s "Follow this event" feature.',
  {
    eventId: z.string().describe('Event ID from get_calendar_events'),
    timezone: z.string().optional().default('W. Europe Standard Time')
      .describe('Timezone for returned event times'),
  },
  async ({ eventId, timezone }) => {
    const event = await calendarClient.followEvent(eventId, timezone);
    return { content: [{ type: 'text', text: JSON.stringify(event, null, 2) }] };
  }
);
```

- [ ] **Step 7: Build and smoke test**

```bash
npm run build 2>&1 | tail -5
```

Verify all tools appear:
```bash
echo '{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2024-11-05","capabilities":{},"clientInfo":{"name":"test","version":"1"}}}
{"jsonrpc":"2.0","id":2,"method":"tools/list","params":{}}' | node dist/index.js 2>/dev/null | jq -r '.result.tools[].name' 2>/dev/null
```

Expected output:
```
get_calendar_events
create_calendar_event
update_calendar_event
cancel_calendar_event
delete_calendar_event
respond_to_calendar_event
follow_calendar_event
```

- [ ] **Step 8: Commit**

```bash
git add src/index.ts
git commit -m "feat: register write tools — create, update, cancel, delete, RSVP, follow"
```

---

## Task 4: Integration Tests

**Files:** Create `tests/calendar-write.test.ts`

- [ ] **Step 1: Write integration tests**

Tests create a real event, update it, RSVP to it, then delete it. This is a sequential integration test that requires a live Edge session.

```typescript
// tests/calendar-write.test.ts
import { TokenManager } from '../src/auth.js';
import { CalendarClient } from '../src/calendar.js';
import type { OwaCreateEventPayload } from '../src/types.js';

describe('CalendarClient write operations', () => {
  let manager: TokenManager;
  let client: CalendarClient;
  let createdEventId: string;

  beforeAll(async () => {
    manager = new TokenManager();
    client = new CalendarClient(manager);
  });

  afterAll(async () => {
    // Cleanup: delete test event if it still exists
    if (createdEventId) {
      try { await client.deleteEvent(createdEventId); } catch { /* ignore */ }
    }
    await manager.close();
  });

  test('creates an event', async () => {
    const payload: OwaCreateEventPayload = {
      Subject: `owa-mcp test event ${Date.now()}`,
      Start: {
        DateTime: new Date(Date.now() + 7 * 24 * 60 * 60 * 1000).toISOString().replace('Z', '').split('.')[0],
        TimeZone: 'UTC',
      },
      End: {
        DateTime: new Date(Date.now() + 7 * 24 * 60 * 60 * 1000 + 30 * 60 * 1000).toISOString().replace('Z', '').split('.')[0],
        TimeZone: 'UTC',
      },
      ShowAs: 'Free',
      Sensitivity: 'Private',
    };
    const event = await client.createEvent(payload);
    expect(event.id).toBeTruthy();
    expect(event.subject).toMatch(/owa-mcp test event/);
    expect(event.showAs).toBe('Free');
    expect(event.isPrivate).toBe(true);
    createdEventId = event.id;
  }, 40_000);

  test('updates the event', async () => {
    const event = await client.updateEvent(createdEventId, {
      Subject: 'owa-mcp updated test event',
      ShowAs: 'Busy',
    });
    expect(event.subject).toBe('owa-mcp updated test event');
    expect(event.showAs).toBe('Busy');
  }, 20_000);

  test('deletes the event', async () => {
    await client.deleteEvent(createdEventId);
    createdEventId = ''; // prevent afterAll double-delete
  }, 20_000);
});
```

Note: RSVP and cancel tests are not included because they require events organized by other users. The create/update/delete flow covers the core API integration.

- [ ] **Step 2: Run the tests**

```bash
npm test -- --testPathPattern=calendar-write 2>&1 | tail -20
```

Expected: 3 tests PASS (create, update, delete).

- [ ] **Step 3: Commit**

```bash
git add tests/calendar-write.test.ts
git commit -m "test: integration tests for calendar write operations"
```

---

## Task 5: Update Documentation

**Files:** Modify `CLAUDE.md`, `README.md`

- [ ] **Step 1: Update CLAUDE.md**

No structural changes needed — the "Adding new tools" section and output convention already cover the new tools. Just verify it's still accurate.

- [ ] **Step 2: Update README.md Available Tools section**

Replace the single `get_calendar_events` tool table with the full list:

```markdown
## Available Tools

### `get_calendar_events`
Returns calendar events in a time range.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `startDateTime` | string | yes | ISO 8601 start |
| `endDateTime` | string | yes | ISO 8601 end |
| `maxResults` | number | no | Max events (default 50, max 100) |
| `timezone` | string | no | IANA timezone (default UTC) |

### `create_calendar_event`
Create a new event. Adding attendees auto-sends invitations.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `subject` | string | yes | Event title |
| `startDateTime` | string | yes | Local datetime without offset |
| `endDateTime` | string | yes | Local datetime without offset |
| `timezone` | string | no | Windows timezone name (default "W. Europe Standard Time") |
| `body` | string | no | Event description |
| `location` | string | no | Location name |
| `attendees` | array | no | `[{ email, name?, type? }]` — sends invitations |
| `isAllDay` | boolean | no | All-day event |
| `showAs` | string | no | Free, Tentative, Busy, Oof, WorkingElsewhere |
| `isOnlineMeeting` | boolean | no | Create as Teams meeting |

### `update_calendar_event`
Update fields on an existing event. Only include fields to change.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `eventId` | string | yes | Event ID |
| `subject` | string | no | New title |
| `startDateTime` | string | no | New start time |
| `endDateTime` | string | no | New end time |
| `timezone` | string | no | Timezone for start/end |
| `body` | string | no | New body (caution: overwrites Teams join link) |
| `location` | string | no | New location |
| `showAs` | string | no | New show-as status |

### `cancel_calendar_event`
Cancel a meeting you organized. Sends cancellation with reason to attendees.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `eventId` | string | yes | Event ID |
| `reason` | string | no | Cancellation reason sent to attendees |

### `delete_calendar_event`
Remove an event from your calendar silently (no notification sent).

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `eventId` | string | yes | Event ID |

### `respond_to_calendar_event`
RSVP to a meeting: accept, tentatively accept, or decline.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `eventId` | string | yes | Event ID |
| `response` | string | yes | `accept`, `tentative`, or `decline` |
| `comment` | string | no | Message to organizer |
| `sendResponse` | boolean | no | Notify organizer (default true) |
| `proposedStartDateTime` | string | no | Propose alternative start (tentative/decline only) |
| `proposedEndDateTime` | string | no | Propose alternative end |

### `follow_calendar_event`
Track an event on your calendar without RSVPing. Shows as Free, organizer not notified.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `eventId` | string | yes | Event ID |
| `timezone` | string | no | Timezone for returned event |
```

- [ ] **Step 3: Update README.md Roadmap**

Replace the roadmap checkboxes — mark calendar items as done, keep mail items:

```markdown
## Roadmap

- [x] `create_calendar_event`
- [x] `update_calendar_event`
- [x] `cancel_calendar_event` (with reason)
- [x] `respond_to_calendar_event` (accept / tentative / decline)
- [x] `follow_calendar_event`
- [x] `delete_calendar_event`
- [ ] `get_emails` (Mail.ReadWrite scope is already present in the token)
- [ ] `send_email`
```

- [ ] **Step 4: Commit**

```bash
git add CLAUDE.md README.md
git commit -m "docs: document all calendar write tools in README"
```

---

## Task 6: Build, Push, Reload

- [ ] **Step 1: Final build**

```bash
npm run build
```

- [ ] **Step 2: Push to GitHub**

```bash
git push
```

- [ ] **Step 3: Reload MCP server**

Run `/mcp` in Claude Code to reconnect. Verify all 7 tools appear.

---

## Verification

After implementation, verify end-to-end:

1. **Run all tests:** `npm test` — all existing + new tests should pass
2. **Build succeeds:** `npm run build` — no errors
3. **Tool list:** All 7 tools appear after `/mcp` reload
4. **Manual smoke test via Claude Code:**
   - Ask "Create a test event tomorrow at 3pm for 30 minutes called 'MCP test'"
   - Ask "Update that event to 3:30pm"
   - Ask "Delete that test event"
5. **Type check:** `npx tsc --noEmit` passes

---

## Design Notes

**Why Windows timezone names?** The OWA v2.0 API requires Windows timezone names in `Start.TimeZone` / `End.TimeZone` (e.g., "W. Europe Standard Time" not "Europe/Berlin"). The `get_calendar_events` tool uses `Prefer: outlook.timezone` header which accepts IANA names for *response* formatting, but write operations need Windows names in the body. We default to "W. Europe Standard Time" since the user is in Germany.

**Why separate cancel vs delete?** Cancel is organizer-only and sends a message to attendees. Delete silently removes from your calendar. Different intents, different API endpoints, different permissions.

**Why one respond tool instead of three?** Accept/tentative/decline share parameters. One tool with a `response` enum is cleaner for the LLM to reason about.

**Follow emulation:** `tentativelyAccept(sendResponse=false)` + `PATCH showAs=Free`. This is the closest API approximation to New Outlook's "Follow" feature, which has no public API. The organizer sees no notification, and the event appears on the user's calendar without blocking time.
