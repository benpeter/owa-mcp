# Meeting Series Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add recurring meeting series awareness to owa-mcp — enrich events with series metadata, extend cancel/delete with series scope, and add tools to inspect series masters and list instances.

**Architecture:** The OWA REST API already returns `Type` and `SeriesMasterId` on calendar events; we just need to request them in `$select` and surface them through `normalise()`. Series operations (cancel/delete with scope, get master, list instances) build on two shared helpers: `resolveSeriesMasterId` and `getSeriesRecurrence`. All changes touch 4 files: types, calendar client, index (tool registration), and a new test file.

**Tech Stack:** TypeScript, OWA REST API v2.0, Zod schemas, Jest (integration tests against live Edge session)

---

## File Structure

| File | Action | Responsibility |
|------|--------|----------------|
| `src/types.ts` | Modify | Add `RecurrencePattern`, `RecurrenceRange`, `RecurrenceInfo` interfaces. Extend `OwaCalendarEvent` with `Type`, `SeriesMasterId`. Extend `CalendarEvent` with `type`, `seriesMasterId`, `recurrence`. |
| `src/calendar.ts` | Modify | Update `$select`, update `normalise()`, add `resolveSeriesMasterId()`, `getSeriesRecurrence()`, `getSeriesMaster()`, `listSeriesInstances()`. Extend `cancelEvent()` and `deleteEvent()` with `scope` parameter. |
| `src/index.ts` | Modify | Update `cancel_calendar_event` and `delete_calendar_event` tool schemas (add `scope`). Register `get_series_master` and `list_series_instances` tools. |
| `tests/calendar-series.test.ts` | Create | Integration tests for all series operations. |

---

### Task 1: Add Recurrence Types to `src/types.ts`

**Files:**
- Modify: `src/types.ts`

- [ ] **Step 1: Add the three recurrence interfaces after `OwaToken`**

Add these interfaces after line 8 (after the `OwaToken` interface closing brace):

```typescript
export interface RecurrencePattern {
  type: string;           // "daily" | "weekly" | "absoluteMonthly" | "relativeMonthly" | "absoluteYearly" | "relativeYearly"
  interval: number;
  daysOfWeek?: string[];
  dayOfMonth?: number;
  month?: number;
  index?: string;
  firstDayOfWeek?: string;
}

export interface RecurrenceRange {
  type: string;           // "endDate" | "numbered" | "noEnd"
  startDate: string;
  endDate?: string;
  numberOfOccurrences?: number;
  recurrenceTimeZone?: string;
}

export interface RecurrenceInfo {
  pattern: RecurrencePattern;
  range: RecurrenceRange;
}
```

- [ ] **Step 2: Extend `CalendarEvent` with series fields**

Add three fields after `isRecurring: boolean;` (line 21) in the `CalendarEvent` interface:

```typescript
  type: 'singleInstance' | 'occurrence' | 'exception' | 'seriesMaster';
  seriesMasterId: string | null;
  recurrence: RecurrenceInfo | null;
```

- [ ] **Step 3: Extend `OwaCalendarEvent` with raw API fields**

Add two fields to `OwaCalendarEvent` after the `Sensitivity` field (line 43):

```typescript
  Type: string;
  SeriesMasterId: string | null;
```

- [ ] **Step 4: Verify the project still compiles**

Run: `cd /Users/I752296/github/benpeter/owa-mcp && npx tsc --noEmit`
Expected: No errors (existing code doesn't yet produce the new fields, but types are additive — the compile will fail because `normalise()` doesn't return the new fields. That's expected and will be fixed in Task 2.)

Actually, since `normalise()` returns `CalendarEvent` but doesn't include the new required fields, this will fail. That's the correct TDD signal — proceed to Task 2.

- [ ] **Step 5: Commit**

```bash
git add src/types.ts
git commit -m "feat(types): add RecurrenceInfo and series fields to CalendarEvent"
```

---

### Task 2: Update `normalise()` and `$select` in `src/calendar.ts`

**Files:**
- Modify: `src/calendar.ts`

- [ ] **Step 1: Add `Type,SeriesMasterId` to the `$select` clause**

In `getCalendarEvents()`, find the `$select` param string (line 34) and add `Type,SeriesMasterId` to it:

```typescript
      '$select': 'Id,Subject,Start,End,IsAllDay,Organizer,Location,IsOnlineMeeting,ShowAs,Recurrence,Sensitivity,BodyPreview,Type,SeriesMasterId',
```

- [ ] **Step 2: Add `RecurrenceInfo` to the import from `types.js`**

Update the import at line 3 to include `RecurrenceInfo`:

```typescript
import type {
  CalendarEvent,
  OwaCalendarViewResponse,
  OwaCalendarEvent,
  OwaCreateEventPayload,
  OwaUpdateEventPayload,
  RsvpAction,
  OwaRsvpPayload,
  RecurrenceInfo,
} from './types.js';
```

- [ ] **Step 3: Update `normalise()` to map the new fields**

Replace the `normalise` method body (lines 276–291) with:

```typescript
  private normalise(raw: OwaCalendarEvent): CalendarEvent {
    const type = (raw.Type?.toLowerCase() ?? 'singleInstance') as CalendarEvent['type'];
    const recurrence = raw.Recurrence
      ? this.normaliseRecurrence(raw.Recurrence)
      : null;

    return {
      id: raw.Id,
      subject: raw.Subject,
      start: raw.Start.DateTime,
      end: raw.End.DateTime,
      isAllDay: raw.IsAllDay,
      organizer: raw.Organizer?.EmailAddress?.Name ?? '',
      location: raw.Location?.DisplayName ?? '',
      isOnlineMeeting: raw.IsOnlineMeeting,
      showAs: raw.ShowAs,
      isRecurring: type !== 'singleInstance',
      isPrivate: raw.Sensitivity === 'Private',
      bodyPreview: raw.BodyPreview ?? '',
      type,
      seriesMasterId: raw.SeriesMasterId ?? null,
      recurrence,
    };
  }

  private normaliseRecurrence(raw: unknown): RecurrenceInfo | null {
    if (!raw || typeof raw !== 'object') return null;
    const rec = raw as Record<string, unknown>;
    const pattern = rec.Pattern as Record<string, unknown> | undefined;
    const range = rec.Range as Record<string, unknown> | undefined;
    if (!pattern || !range) return null;

    return {
      pattern: {
        type: String(pattern.Type ?? '').toLowerCase(),
        interval: Number(pattern.Interval ?? 1),
        daysOfWeek: Array.isArray(pattern.DaysOfWeek)
          ? pattern.DaysOfWeek.map((d: unknown) => String(d).toLowerCase())
          : undefined,
        dayOfMonth: pattern.DayOfMonth != null ? Number(pattern.DayOfMonth) : undefined,
        month: pattern.Month != null ? Number(pattern.Month) : undefined,
        index: pattern.Index != null ? String(pattern.Index).toLowerCase() : undefined,
        firstDayOfWeek: pattern.FirstDayOfWeek != null
          ? String(pattern.FirstDayOfWeek).toLowerCase()
          : undefined,
      },
      range: {
        type: String(range.Type ?? '').toLowerCase(),
        startDate: String(range.StartDate ?? ''),
        endDate: range.EndDate != null ? String(range.EndDate) : undefined,
        numberOfOccurrences: range.NumberOfOccurrences != null
          ? Number(range.NumberOfOccurrences)
          : undefined,
        recurrenceTimeZone: range.RecurrenceTimeZone != null
          ? String(range.RecurrenceTimeZone)
          : undefined,
      },
    };
  }
```

- [ ] **Step 4: Verify the project compiles**

Run: `cd /Users/I752296/github/benpeter/owa-mcp && npx tsc --noEmit`
Expected: No errors.

- [ ] **Step 5: Commit**

```bash
git add src/calendar.ts
git commit -m "feat(calendar): add series fields to normalise() and \$select"
```

---

### Task 3: Write Integration Tests for Series Metadata

**Files:**
- Create: `tests/calendar-series.test.ts`

- [ ] **Step 1: Write the test file with metadata tests**

Create `tests/calendar-series.test.ts`:

```typescript
// tests/calendar-series.test.ts
import { TokenManager } from '../src/auth.js';
import { CalendarClient } from '../src/calendar.js';
import type { OwaCreateEventPayload, CalendarEvent } from '../src/types.js';

describe('CalendarClient series operations', () => {
  let manager: TokenManager;
  let client: CalendarClient;

  beforeAll(async () => {
    manager = new TokenManager();
    client = new CalendarClient(manager);
  });

  afterAll(async () => {
    await manager.close();
  });

  describe('series metadata in calendarview', () => {
    test('events include type field', async () => {
      const start = new Date();
      const end = new Date(Date.now() + 14 * 24 * 60 * 60 * 1000);
      const events = await client.getCalendarEvents(start.toISOString(), end.toISOString());

      expect(events.length).toBeGreaterThan(0);
      for (const event of events) {
        expect(['singleInstance', 'occurrence', 'exception', 'seriesMaster']).toContain(event.type);
        expect(event).toHaveProperty('seriesMasterId');
        expect(event).toHaveProperty('recurrence');
      }
    }, 40_000);

    test('recurring occurrences have seriesMasterId set', async () => {
      const start = new Date();
      const end = new Date(Date.now() + 14 * 24 * 60 * 60 * 1000);
      const events = await client.getCalendarEvents(start.toISOString(), end.toISOString());

      const occurrences = events.filter(e => e.type === 'occurrence');
      if (occurrences.length === 0) {
        console.warn('No recurring occurrences found in date range — skipping seriesMasterId assertion');
        return;
      }
      for (const occ of occurrences) {
        expect(occ.seriesMasterId).toBeTruthy();
        expect(occ.isRecurring).toBe(true);
      }
    }, 40_000);

    test('single instances have null seriesMasterId and null recurrence', async () => {
      const start = new Date();
      const end = new Date(Date.now() + 14 * 24 * 60 * 60 * 1000);
      const events = await client.getCalendarEvents(start.toISOString(), end.toISOString());

      const singles = events.filter(e => e.type === 'singleInstance');
      if (singles.length === 0) {
        console.warn('No single-instance events found — skipping assertion');
        return;
      }
      for (const s of singles) {
        expect(s.seriesMasterId).toBeNull();
        expect(s.recurrence).toBeNull();
        expect(s.isRecurring).toBe(false);
      }
    }, 40_000);
  });
});
```

- [ ] **Step 2: Run the metadata tests**

Run: `cd /Users/I752296/github/benpeter/owa-mcp && npm test -- --testPathPattern=calendar-series`
Expected: All 3 tests PASS.

- [ ] **Step 3: Commit**

```bash
git add tests/calendar-series.test.ts
git commit -m "test: add series metadata integration tests"
```

---

### Task 4: Add `resolveSeriesMasterId` and `getSeriesRecurrence` Helpers

**Files:**
- Modify: `src/calendar.ts`

- [ ] **Step 1: Add `resolveSeriesMasterId` method to `CalendarClient`**

Add this method after `followEvent()` and before the `private toServiceId()` method:

```typescript
  /**
   * Resolve any event ID to its series master ID.
   * Returns the event's own ID if it is already a series master.
   * Throws if the event is a singleInstance (not part of a series).
   */
  async resolveSeriesMasterId(eventId: string): Promise<string> {
    const res = await this.request('GET', `/me/events/${eventId}?$select=Id,Type,SeriesMasterId`);
    const raw = (await res.json()) as { Id: string; Type: string; SeriesMasterId: string | null };
    const type = raw.Type?.toLowerCase();

    if (type === 'singleinstance') {
      throw new Error('This event is not part of a recurring series');
    }
    if (type === 'seriesmaster') {
      return raw.Id;
    }
    // occurrence or exception
    if (!raw.SeriesMasterId) {
      throw new Error(`Event ${eventId} is ${type} but has no SeriesMasterId`);
    }
    return raw.SeriesMasterId;
  }

  /**
   * Fetch the series master event and return its current recurrence pattern.
   * Used by thisAndFollowing scope to construct the PATCH payload.
   */
  async getSeriesRecurrence(seriesMasterId: string): Promise<{ recurrence: unknown; event: OwaCalendarEvent }> {
    const res = await this.request('GET', `/me/events/${seriesMasterId}?$select=Id,Subject,Start,End,Recurrence,Type,IsAllDay,Organizer,Location,IsOnlineMeeting,ShowAs,Sensitivity,BodyPreview,SeriesMasterId`);
    const raw = (await res.json()) as OwaCalendarEvent & { Recurrence: unknown };
    if (!raw.Recurrence) {
      throw new Error(`Series master ${seriesMasterId} has no Recurrence`);
    }
    return { recurrence: raw.Recurrence, event: raw };
  }
```

- [ ] **Step 2: Verify the project compiles**

Run: `cd /Users/I752296/github/benpeter/owa-mcp && npx tsc --noEmit`
Expected: No errors.

- [ ] **Step 3: Add integration tests for `resolveSeriesMasterId`**

Append to the `describe('CalendarClient series operations', ...)` block in `tests/calendar-series.test.ts`, after the closing `});` of the `describe('series metadata in calendarview', ...)` block:

```typescript
  describe('resolveSeriesMasterId', () => {
    test('resolves occurrence to series master ID', async () => {
      const start = new Date();
      const end = new Date(Date.now() + 14 * 24 * 60 * 60 * 1000);
      const events = await client.getCalendarEvents(start.toISOString(), end.toISOString());
      const occurrence = events.find(e => e.type === 'occurrence');
      if (!occurrence) {
        console.warn('No recurring occurrence found — skipping resolveSeriesMasterId test');
        return;
      }

      const masterId = await client.resolveSeriesMasterId(occurrence.id);
      expect(masterId).toBeTruthy();
      expect(masterId).toBe(occurrence.seriesMasterId);
    }, 40_000);

    test('returns own ID for series master', async () => {
      const start = new Date();
      const end = new Date(Date.now() + 14 * 24 * 60 * 60 * 1000);
      const events = await client.getCalendarEvents(start.toISOString(), end.toISOString());
      const occurrence = events.find(e => e.type === 'occurrence' && e.seriesMasterId);
      if (!occurrence) {
        console.warn('No recurring occurrence found — skipping series master resolution test');
        return;
      }

      // resolveSeriesMasterId on the master itself should return the same ID
      const masterId = await client.resolveSeriesMasterId(occurrence.seriesMasterId!);
      expect(masterId).toBe(occurrence.seriesMasterId);
    }, 40_000);

    test('throws for single instance events', async () => {
      const start = new Date();
      const end = new Date(Date.now() + 14 * 24 * 60 * 60 * 1000);
      const events = await client.getCalendarEvents(start.toISOString(), end.toISOString());
      const single = events.find(e => e.type === 'singleInstance');
      if (!single) {
        console.warn('No single-instance event found — skipping test');
        return;
      }

      await expect(client.resolveSeriesMasterId(single.id))
        .rejects.toThrow('not part of a recurring series');
    }, 40_000);
  });
```

- [ ] **Step 4: Run tests**

Run: `cd /Users/I752296/github/benpeter/owa-mcp && npm test -- --testPathPattern=calendar-series`
Expected: All tests PASS.

- [ ] **Step 5: Commit**

```bash
git add src/calendar.ts tests/calendar-series.test.ts
git commit -m "feat(calendar): add resolveSeriesMasterId and getSeriesRecurrence helpers"
```

---

### Task 5: Add `getSeriesMaster` Method and Tool

**Files:**
- Modify: `src/calendar.ts`
- Modify: `src/index.ts`

- [ ] **Step 1: Add `getSeriesMaster()` method to `CalendarClient`**

Add this method after `getSeriesRecurrence()`:

```typescript
  /**
   * Get the series master event for any event in a recurring series.
   * Returns the master with full recurrence info and cancelled occurrences.
   */
  async getSeriesMaster(eventId: string, timezone?: string): Promise<CalendarEvent & { cancelledOccurrences: string[] }> {
    const masterId = await this.resolveSeriesMasterId(eventId);
    const res = await this.request('GET',
      `/me/events/${masterId}?$select=Id,Subject,Start,End,Recurrence,Type,IsAllDay,Organizer,Location,IsOnlineMeeting,ShowAs,Sensitivity,BodyPreview,SeriesMasterId,CancelledOccurrences`,
      { timezone }
    );
    const raw = (await res.json()) as OwaCalendarEvent & { CancelledOccurrences?: { Start: { DateTime: string } }[] };
    const normalised = this.normalise(raw);
    const cancelledOccurrences = (raw.CancelledOccurrences ?? []).map(
      (c: { Start: { DateTime: string } }) => c.Start.DateTime
    );
    return { ...normalised, cancelledOccurrences };
  }
```

- [ ] **Step 2: Register the `get_series_master` tool in `src/index.ts`**

Add after the `follow_calendar_event` tool registration (after line 224):

```typescript
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
```

- [ ] **Step 3: Add integration test for `getSeriesMaster`**

Append to the outer `describe('CalendarClient series operations', ...)` block in `tests/calendar-series.test.ts`:

```typescript
  describe('getSeriesMaster', () => {
    test('returns master with recurrence from occurrence ID', async () => {
      const start = new Date();
      const end = new Date(Date.now() + 14 * 24 * 60 * 60 * 1000);
      const events = await client.getCalendarEvents(start.toISOString(), end.toISOString());
      const occurrence = events.find(e => e.type === 'occurrence');
      if (!occurrence) {
        console.warn('No recurring occurrence found — skipping getSeriesMaster test');
        return;
      }

      const master = await client.getSeriesMaster(occurrence.id);
      expect(master.type).toBe('seriesmaster');
      expect(master.id).toBe(occurrence.seriesMasterId);
      expect(master.recurrence).not.toBeNull();
      expect(master.recurrence!.pattern).toHaveProperty('type');
      expect(master.recurrence!.pattern).toHaveProperty('interval');
      expect(master.recurrence!.range).toHaveProperty('type');
      expect(master.recurrence!.range).toHaveProperty('startDate');
      expect(master).toHaveProperty('cancelledOccurrences');
      expect(Array.isArray(master.cancelledOccurrences)).toBe(true);
    }, 40_000);
  });
```

- [ ] **Step 4: Run tests**

Run: `cd /Users/I752296/github/benpeter/owa-mcp && npm test -- --testPathPattern=calendar-series`
Expected: All tests PASS.

- [ ] **Step 5: Commit**

```bash
git add src/calendar.ts src/index.ts tests/calendar-series.test.ts
git commit -m "feat: add get_series_master tool"
```

---

### Task 6: Add `listSeriesInstances` Method and Tool

**Files:**
- Modify: `src/calendar.ts`
- Modify: `src/index.ts`

- [ ] **Step 1: Add `listSeriesInstances()` method to `CalendarClient`**

Add after `getSeriesMaster()`:

```typescript
  /**
   * List all occurrences of a recurring series within a date range.
   * Accepts any event ID from the series (resolved to master automatically).
   */
  async listSeriesInstances(
    eventId: string,
    startDateTime: string,
    endDateTime: string,
    timezone?: string
  ): Promise<CalendarEvent[]> {
    const masterId = await this.resolveSeriesMasterId(eventId);
    const params = new URLSearchParams({ startDateTime, endDateTime });
    const res = await this.request('GET',
      `/me/events/${masterId}/instances?${params}`,
      { timezone }
    );
    const data = (await res.json()) as OwaCalendarViewResponse;
    return data.value.map(raw => this.normalise(raw));
  }
```

- [ ] **Step 2: Register the `list_series_instances` tool in `src/index.ts`**

Add after the `get_series_master` tool registration:

```typescript
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
```

- [ ] **Step 3: Add integration test for `listSeriesInstances`**

Append to the outer `describe` block in `tests/calendar-series.test.ts`:

```typescript
  describe('listSeriesInstances', () => {
    test('returns instances for a recurring series', async () => {
      const start = new Date();
      const end = new Date(Date.now() + 14 * 24 * 60 * 60 * 1000);
      const events = await client.getCalendarEvents(start.toISOString(), end.toISOString());
      const occurrence = events.find(e => e.type === 'occurrence');
      if (!occurrence) {
        console.warn('No recurring occurrence found — skipping listSeriesInstances test');
        return;
      }

      // List instances over a 30-day window
      const rangeStart = new Date();
      const rangeEnd = new Date(Date.now() + 30 * 24 * 60 * 60 * 1000);
      const instances = await client.listSeriesInstances(
        occurrence.id,
        rangeStart.toISOString(),
        rangeEnd.toISOString()
      );

      expect(Array.isArray(instances)).toBe(true);
      expect(instances.length).toBeGreaterThan(0);
      for (const inst of instances) {
        expect(['occurrence', 'exception']).toContain(inst.type);
        expect(inst.seriesMasterId).toBe(occurrence.seriesMasterId);
      }
    }, 40_000);
  });
```

- [ ] **Step 4: Run tests**

Run: `cd /Users/I752296/github/benpeter/owa-mcp && npm test -- --testPathPattern=calendar-series`
Expected: All tests PASS.

- [ ] **Step 5: Commit**

```bash
git add src/calendar.ts src/index.ts tests/calendar-series.test.ts
git commit -m "feat: add list_series_instances tool"
```

---

### Task 7: Extend `cancelEvent` with `scope` Parameter

**Files:**
- Modify: `src/calendar.ts`
- Modify: `src/index.ts`

- [ ] **Step 1: Update `cancelEvent()` method signature and implementation**

Replace the current `cancelEvent` method (line 74–78) with:

```typescript
  async cancelEvent(eventId: string, comment?: string, scope: 'single' | 'thisAndFollowing' | 'allInSeries' = 'single'): Promise<void> {
    if (scope === 'single') {
      await this.request('POST', `/me/events/${eventId}/cancel`, {
        body: comment ? { Comment: comment } : {},
      });
      return;
    }

    if (scope === 'allInSeries') {
      const masterId = await this.resolveSeriesMasterId(eventId);
      await this.request('POST', `/me/events/${masterId}/cancel`, {
        body: comment ? { Comment: comment } : {},
      });
      return;
    }

    // thisAndFollowing: truncate the series recurrence range
    const masterId = await this.resolveSeriesMasterId(eventId);

    // Get the occurrence's start date to compute the new end date
    const occRes = await this.request('GET', `/me/events/${eventId}?$select=Start,Type`);
    const occ = (await occRes.json()) as { Start: { DateTime: string }; Type: string };
    if (occ.Type?.toLowerCase() === 'seriesmaster') {
      throw new Error('Cannot use thisAndFollowing on the series master itself — use allInSeries instead');
    }

    // Fetch current recurrence from the master
    const { recurrence } = await this.getSeriesRecurrence(masterId);
    const rec = recurrence as { Pattern: Record<string, unknown>; Range: Record<string, unknown> };

    // New end date = occurrence start date minus one day (date-only, no timezone conversion)
    const occDate = occ.Start.DateTime.split('T')[0]; // "2026-04-15"
    const endDate = new Date(occDate + 'T00:00:00Z');
    endDate.setUTCDate(endDate.getUTCDate() - 1);
    const newEndDate = endDate.toISOString().split('T')[0]; // "2026-04-14"

    await this.request('PATCH', `/me/events/${masterId}`, {
      body: {
        Recurrence: {
          Pattern: rec.Pattern,
          Range: {
            ...rec.Range,
            Type: 'EndDate',
            EndDate: newEndDate,
          },
        },
      },
    });
  }
```

- [ ] **Step 2: Update `cancel_calendar_event` tool schema in `src/index.ts`**

Replace the `cancel_calendar_event` tool registration (lines 147–158) with:

```typescript
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
```

- [ ] **Step 3: Verify the project compiles**

Run: `cd /Users/I752296/github/benpeter/owa-mcp && npx tsc --noEmit`
Expected: No errors.

- [ ] **Step 4: Commit**

```bash
git add src/calendar.ts src/index.ts
git commit -m "feat: extend cancel_calendar_event with scope parameter"
```

---

### Task 8: Extend `deleteEvent` with `scope` Parameter

**Files:**
- Modify: `src/calendar.ts`
- Modify: `src/index.ts`

- [ ] **Step 1: Update `deleteEvent()` method signature and implementation**

Replace the current `deleteEvent` method (line 80–82, though line numbers will have shifted) with:

```typescript
  async deleteEvent(eventId: string, scope: 'single' | 'thisAndFollowing' | 'allInSeries' = 'single'): Promise<void> {
    if (scope === 'single') {
      await this.request('DELETE', `/me/events/${eventId}`);
      return;
    }

    if (scope === 'allInSeries') {
      const masterId = await this.resolveSeriesMasterId(eventId);
      await this.request('DELETE', `/me/events/${masterId}`);
      return;
    }

    // thisAndFollowing: same truncation approach as cancelEvent
    const masterId = await this.resolveSeriesMasterId(eventId);

    const occRes = await this.request('GET', `/me/events/${eventId}?$select=Start,Type`);
    const occ = (await occRes.json()) as { Start: { DateTime: string }; Type: string };
    if (occ.Type?.toLowerCase() === 'seriesmaster') {
      throw new Error('Cannot use thisAndFollowing on the series master itself — use allInSeries instead');
    }

    const { recurrence } = await this.getSeriesRecurrence(masterId);
    const rec = recurrence as { Pattern: Record<string, unknown>; Range: Record<string, unknown> };

    const occDate = occ.Start.DateTime.split('T')[0];
    const endDate = new Date(occDate + 'T00:00:00Z');
    endDate.setUTCDate(endDate.getUTCDate() - 1);
    const newEndDate = endDate.toISOString().split('T')[0];

    await this.request('PATCH', `/me/events/${masterId}`, {
      body: {
        Recurrence: {
          Pattern: rec.Pattern,
          Range: {
            ...rec.Range,
            Type: 'EndDate',
            EndDate: newEndDate,
          },
        },
      },
    });
  }
```

- [ ] **Step 2: Update `delete_calendar_event` tool schema in `src/index.ts`**

Replace the `delete_calendar_event` tool registration with:

```typescript
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
```

- [ ] **Step 3: Update the `afterAll` cleanup in `tests/calendar-write.test.ts`**

The `afterAll` in `tests/calendar-write.test.ts` calls `client.deleteEvent(createdEventId)`. Since `deleteEvent` now takes an optional second parameter, this call is still valid (defaults to `'single'`). No change needed.

- [ ] **Step 4: Verify the project compiles**

Run: `cd /Users/I752296/github/benpeter/owa-mcp && npx tsc --noEmit`
Expected: No errors.

- [ ] **Step 5: Commit**

```bash
git add src/calendar.ts src/index.ts
git commit -m "feat: extend delete_calendar_event with scope parameter"
```

---

### Task 9: Integration Tests for Series Cancel/Delete Operations

**Files:**
- Modify: `tests/calendar-series.test.ts`

These tests create a temporary weekly recurring event, test series operations against it, and clean up.

- [ ] **Step 1: Add series write operation tests**

Append to the outer `describe` block in `tests/calendar-series.test.ts`:

```typescript
  describe('series cancel/delete operations', () => {
    let seriesEventId: string;

    afterEach(async () => {
      // Clean up any test series that wasn't already deleted
      if (seriesEventId) {
        try { await client.deleteEvent(seriesEventId, 'allInSeries'); } catch { /* ignore */ }
        seriesEventId = '';
      }
    });

    async function createTestSeries(): Promise<CalendarEvent> {
      const startDate = new Date(Date.now() + 7 * 24 * 60 * 60 * 1000);
      const startStr = startDate.toISOString().replace('Z', '').split('.')[0];
      const endDate = new Date(startDate.getTime() + 30 * 60 * 1000);
      const endStr = endDate.toISOString().replace('Z', '').split('.')[0];

      // Create a weekly recurring event for 5 weeks
      const payload: OwaCreateEventPayload = {
        Subject: `owa-mcp series test ${Date.now()}`,
        Start: { DateTime: startStr, TimeZone: 'UTC' },
        End: { DateTime: endStr, TimeZone: 'UTC' },
        ShowAs: 'Free',
        Sensitivity: 'Private',
        Recurrence: {
          Pattern: {
            Type: 'Weekly',
            Interval: 1,
            DaysOfWeek: [startDate.toLocaleDateString('en-US', { weekday: 'long' })],
            FirstDayOfWeek: 'Sunday',
          },
          Range: {
            Type: 'Numbered',
            StartDate: startStr.split('T')[0],
            NumberOfOccurrences: 5,
          },
        },
      };
      const event = await client.createEvent(payload);
      seriesEventId = event.id;
      return event;
    }

    test('cancel allInSeries removes the entire series', async () => {
      const created = await createTestSeries();

      // Cancel all
      await client.cancelEvent(created.id, 'Test cleanup', 'allInSeries');

      // Verify: listing instances over the next 60 days should return nothing or throw
      // The master itself is cancelled, so resolve should fail or return empty
      const rangeStart = new Date();
      const rangeEnd = new Date(Date.now() + 60 * 24 * 60 * 60 * 1000);

      try {
        const instances = await client.listSeriesInstances(
          created.id,
          rangeStart.toISOString(),
          rangeEnd.toISOString()
        );
        // If we get results, they should be empty (series was cancelled)
        expect(instances.length).toBe(0);
      } catch (err) {
        // API may return 404 for cancelled series — that's also acceptable
        expect(String(err)).toMatch(/404|not found|ErrorItemNotFound/i);
      }
      seriesEventId = ''; // Already cleaned up via cancel
    }, 60_000);

    test('delete allInSeries removes the entire series', async () => {
      const created = await createTestSeries();

      await client.deleteEvent(created.id, 'allInSeries');

      const rangeStart = new Date();
      const rangeEnd = new Date(Date.now() + 60 * 24 * 60 * 60 * 1000);

      try {
        const instances = await client.listSeriesInstances(
          created.id,
          rangeStart.toISOString(),
          rangeEnd.toISOString()
        );
        expect(instances.length).toBe(0);
      } catch (err) {
        expect(String(err)).toMatch(/404|not found|ErrorItemNotFound/i);
      }
      seriesEventId = ''; // Already cleaned up
    }, 60_000);

    test('thisAndFollowing truncates the series', async () => {
      const created = await createTestSeries();

      // List instances to find the third occurrence
      const rangeStart = new Date();
      const rangeEnd = new Date(Date.now() + 60 * 24 * 60 * 60 * 1000);
      const instances = await client.listSeriesInstances(
        created.id,
        rangeStart.toISOString(),
        rangeEnd.toISOString()
      );
      expect(instances.length).toBe(5);

      // Cancel from the 3rd occurrence onward
      const thirdInstance = instances[2];
      await client.cancelEvent(thirdInstance.id, undefined, 'thisAndFollowing');

      // Re-list: should now have only 2 instances
      const afterInstances = await client.listSeriesInstances(
        created.id,
        rangeStart.toISOString(),
        rangeEnd.toISOString()
      );
      expect(afterInstances.length).toBe(2);
    }, 60_000);
  });
```

- [ ] **Step 2: Run all series tests**

Run: `cd /Users/I752296/github/benpeter/owa-mcp && npm test -- --testPathPattern=calendar-series`
Expected: All tests PASS.

- [ ] **Step 3: Commit**

```bash
git add tests/calendar-series.test.ts
git commit -m "test: add series cancel/delete integration tests"
```

---

### Task 10: Run Full Test Suite and Final Verification

**Files:** None (verification only)

- [ ] **Step 1: Run the full test suite**

Run: `cd /Users/I752296/github/benpeter/owa-mcp && npm test`
Expected: All tests pass — existing calendar, calendar-write, mail, and auth tests should be unaffected by the changes.

- [ ] **Step 2: Verify build succeeds**

Run: `cd /Users/I752296/github/benpeter/owa-mcp && npm run build`
Expected: Build completes with no errors, output in `dist/`.

- [ ] **Step 3: Verify type checking**

Run: `cd /Users/I752296/github/benpeter/owa-mcp && npx tsc --noEmit`
Expected: No type errors.

- [ ] **Step 4: Commit any remaining changes**

If there are any uncommitted fixes from the verification steps:

```bash
git add -A
git commit -m "fix: address issues found during verification"
```

---

## Notes for the Implementer

1. **Live Edge session required**: All tests are integration tests. You need a signed-in Edge browser with access to outlook.office.com. First run takes ~10s for token acquisition.

2. **Test data dependency**: The metadata tests (Task 3) assume at least one recurring meeting exists in the next 14 days. The write tests (Task 9) create their own temporary recurring events and clean up.

3. **The `thisAndFollowing` date math**: The new end date is `occurrence.Start.DateTime` date portion minus one day. This uses date-only strings (no timezone conversion), consistent with how `Recurrence.Range` works in the OWA API.

4. **`normaliseRecurrence` defensiveness**: The OWA API returns `Recurrence` as `unknown` on the TypeScript side. The normaliser handles missing/malformed data gracefully by returning `null`.

5. **Backward compatibility**: `cancelEvent(eventId, comment)` and `deleteEvent(eventId)` continue to work unchanged — `scope` defaults to `'single'`.
