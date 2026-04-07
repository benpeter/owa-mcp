# Meeting Series Handling & Cancel-From-Date

## Summary

Add recurring meeting series awareness to the owa-mcp server: enrich calendar event responses with series metadata, extend cancel/delete tools with a `scope` parameter for series operations, and add tools to inspect and list series instances.

## Motivation

The server currently treats all calendar events as flat objects. When `calendarview` returns expanded occurrences of recurring series, the series relationship is lost — there's no way to cancel future occurrences, inspect recurrence patterns, or operate on a series as a whole. This forces the LLM to treat every occurrence independently, which doesn't match how users think about recurring meetings.

## Design

### 1. CalendarEvent Type Enrichment

**New fields on `CalendarEvent`** (normalized output):

| Field | Type | Description |
|-------|------|-------------|
| `type` | `"singleInstance" \| "occurrence" \| "exception" \| "seriesMaster"` | Event type (lowercase, mapped from OWA PascalCase) |
| `seriesMasterId` | `string \| null` | Series master event ID (null for singleInstance and seriesMaster) |
| `recurrence` | `RecurrenceInfo \| null` | Full recurrence pattern + range (non-null only on seriesMaster) |

**New types in `src/types.ts`:**

```typescript
interface RecurrencePattern {
  type: string;           // "daily" | "weekly" | "absoluteMonthly" | "relativeMonthly" | "absoluteYearly" | "relativeYearly"
  interval: number;
  daysOfWeek?: string[];
  dayOfMonth?: number;
  month?: number;
  index?: string;
  firstDayOfWeek?: string;
}

interface RecurrenceRange {
  type: string;           // "endDate" | "numbered" | "noEnd"
  startDate: string;
  endDate?: string;
  numberOfOccurrences?: number;
  recurrenceTimeZone?: string;
}

interface RecurrenceInfo {
  pattern: RecurrencePattern;
  range: RecurrenceRange;
}
```

**Changes to `OwaCalendarEvent`:** Add `Type: string` and `SeriesMasterId: string | null`.

**Changes to `$select`:** Add `Type,SeriesMasterId` to the calendarview select clause.

**Changes to `normalise()`:** Map `Type` to lowercase, pass through `SeriesMasterId`, normalize `Recurrence` from `unknown` to `RecurrenceInfo | null`. The existing `isRecurring` field is now derived as `type !== "singleInstance"` (fixing the current bug where occurrences were marked as non-recurring because their `Recurrence` field is null).

### 2. Extended cancel_calendar_event

New parameter:

```
scope: "single" | "thisAndFollowing" | "allInSeries"  (default: "single")
```

**`single`** (default, backward compatible): `POST /me/events/{eventId}/cancel` — current behavior unchanged.

**`thisAndFollowing`**: Truncates the series at the target occurrence.

1. `GET /me/events/{eventId}` — read occurrence, get `SeriesMasterId` and `Start.DateTime`
2. `GET /me/events/{seriesMasterId}` — read current `Recurrence` pattern
3. Calculate new end date: take the target occurrence's `Start.DateTime`, extract the date portion (YYYY-MM-DD), subtract one day. This is a calendar date (no timezone conversion needed — `Recurrence.Range` uses date-only strings like `"2026-04-06"`)
4. `PATCH /me/events/{seriesMasterId}` with full `Recurrence` object, modified `Range.EndDate` and `Range.Type` set to `"endDate"`

Note: PATCH truncation does not send cancellation messages to attendees. Future occurrences silently disappear from attendee calendars.

Validation: Rejects `singleInstance` with clear error "This event is not part of a recurring series".

**`allInSeries`**: Cancels the entire series.

1. If event is `occurrence`/`exception`, resolve `SeriesMasterId` via `GET /me/events/{eventId}`
2. `POST /me/events/{seriesMasterId}/cancel` with the comment

Works whether the caller passes an occurrence ID or a series master ID.

### 3. Extended delete_calendar_event

Same `scope` parameter as cancel:

```
scope: "single" | "thisAndFollowing" | "allInSeries"  (default: "single")
```

- **`single`**: `DELETE /me/events/{eventId}` — current behavior
- **`thisAndFollowing`**: Same PATCH approach as cancel (truncate recurrence range)
- **`allInSeries`**: `DELETE /me/events/{seriesMasterId}` — resolves master first if needed

### 4. New Tool: get_series_master

Inspect the master event of a recurring series.

**Parameters:**
- `eventId` (required): Any event ID from the series — occurrence, exception, or series master. Resolved automatically.

**Implementation:**
1. `GET /me/events/{eventId}` — check `Type`
2. If `occurrence`/`exception`, follow `SeriesMasterId`
3. `GET /me/events/{seriesMasterId}?$select=Id,Subject,Start,End,Recurrence,Type,CancelledOccurrences,Sensitivity,Organizer,Location,IsOnlineMeeting,ShowAs,BodyPreview,SeriesMasterId`
4. Return normalized `CalendarEvent` with full `Recurrence` and `cancelledOccurrences: string[]`

### 5. New Tool: list_series_instances

List all occurrences of a recurring series within a date range.

**Parameters:**
- `eventId` (required): Any event ID from the series (resolved to master automatically)
- `startDateTime` (required): ISO 8601
- `endDateTime` (required): ISO 8601
- `timezone` (optional, default UTC)

**Implementation:**
1. Resolve to series master ID
2. `GET /me/events/{seriesMasterId}/instances?startDateTime=...&endDateTime=...`
3. Return array of normalized `CalendarEvent` objects

### 6. Internal Helpers

**`resolveSeriesMasterId(eventId: string): Promise<string>`**

Shared by cancel, delete, get_series_master, and list_series_instances. Fetches the event, returns:
- `SeriesMasterId` if the event is an occurrence/exception
- The event's own `Id` if it's already a series master
- Throws if `singleInstance` and the caller requires a series context

**`getSeriesRecurrence(seriesMasterId: string): Promise<{recurrence: RecurrenceInfo, event: OwaCalendarEvent}>`**

Used by `thisAndFollowing` scope in cancel/delete. Fetches the series master and returns its current recurrence pattern (needed to construct the complete PATCH payload).

## Files Changed

| File | Change |
|------|--------|
| `src/types.ts` | Add `RecurrencePattern`, `RecurrenceRange`, `RecurrenceInfo` types. Add `Type`, `SeriesMasterId` to `OwaCalendarEvent`. Add `type`, `seriesMasterId`, `recurrence` to `CalendarEvent`. |
| `src/calendar.ts` | Add `Type,SeriesMasterId` to `$select`. Update `normalise()`. Add `resolveSeriesMasterId()`, `getSeriesRecurrence()`, `getSeriesMaster()`, `listSeriesInstances()`. Extend `cancelEvent()` and `deleteEvent()` with scope logic. |
| `src/index.ts` | Update `cancel_calendar_event` and `delete_calendar_event` schemas (add `scope`). Register `get_series_master` and `list_series_instances` tools. |
| `tests/calendar-series.test.ts` | New integration test file for series operations. |

## Testing

Integration tests in `tests/calendar-series.test.ts`:

1. **Series metadata in calendarview** — fetch events containing a known recurring meeting, verify `type`, `seriesMasterId`, `recurrence` fields
2. **get_series_master** — pass an occurrence ID, verify it returns the master with recurrence pattern
3. **list_series_instances** — pass a series event, verify expected occurrences in date range
4. **Cancel single occurrence** — cancel one occurrence, verify it's removed from calendarview
5. **Cancel thisAndFollowing** — create a weekly test series, cancel from a date, verify series end date changed
6. **Cancel allInSeries** — cancel entire test series, verify all occurrences gone

Tests 4-6 create a temporary recurring event and clean up after.

## API Reference

| Operation | Method | OWA Endpoint | ID Used |
|-----------|--------|-------------|---------|
| Fetch expanded occurrences | `GET` | `/me/calendarview?startDateTime=...&endDateTime=...` | N/A |
| Get series master | `GET` | `/me/events/{seriesMasterId}` | Series master ID |
| List series instances | `GET` | `/me/events/{seriesMasterId}/instances?startDateTime=...&endDateTime=...` | Series master ID |
| Cancel single occurrence | `POST` | `/me/events/{occurrenceId}/cancel` | Occurrence ID |
| Cancel this+following | `PATCH` | `/me/events/{seriesMasterId}` (truncate recurrence range) | Series master ID |
| Cancel entire series | `POST` | `/me/events/{seriesMasterId}/cancel` | Series master ID |
| Delete single occurrence | `DELETE` | `/me/events/{occurrenceId}` | Occurrence ID |
| Delete this+following | `PATCH` | `/me/events/{seriesMasterId}` (truncate recurrence range) | Series master ID |
| Delete entire series | `DELETE` | `/me/events/{seriesMasterId}` | Series master ID |

All endpoints use base URL `https://outlook.office.com/api/v2.0`.
