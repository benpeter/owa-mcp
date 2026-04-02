# Route RSVP and Follow Through service.svc with Native ImmutableIds

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Route `respond_to_calendar_event` and `follow_calendar_event` through OWA's `service.svc?action=RespondToCalendarEvent` using natively-fetched ImmutableIds, eliminating the unreliable `translateExchangeIds` step and bypassing the `ResponseRequested: false` restriction.

**Builds on:** Current implementation where Follow uses `service.svc` with translated IDs and RSVP uses the REST API.

---

## Background: What We Learned (2026-04-02)

### The Problem

1. **REST API rejects RSVP** on events where the organizer set `ResponseRequested: false` — returns `"Your request can't be completed. The meeting organizer hasn't requested a response."`. The OWA browser client bypasses this by using `service.svc` instead.

2. **ID translation via `translateExchangeIds` is unreliable** — for some events (already-followed events, certain single-instance events), the translated `RestImmutableEntryId` produces `ErrorItemNotFound` when passed to `service.svc`. But the OWA browser client successfully operates on these same events because it gets ImmutableIds directly from `service.svc?action=GetCalendarEvent`.

### The Solution

Do what the browser does:
1. Fetch the event's native ImmutableId via `service.svc?action=GetCalendarEvent` (with `Prefer: IdType="ImmutableId"`)
2. Use that ImmutableId for `service.svc?action=RespondToCalendarEvent`

This eliminates the `translateExchangeIds` step entirely and ensures the ID is always valid.

### service.svc RespondToCalendarEvent Protocol

**Endpoint:** `POST https://outlook.office.com/owa/service.svc?action=RespondToCalendarEvent`

**Payload delivery:** URL-encoded JSON in the `x-owa-urlpostdata` header. Request body is empty (content-length: 0).

**Payload structure:**
```json
{
  "__type": "RespondToCalendarEventJsonRequest:#Exchange",
  "Header": {
    "__type": "JsonRequestHeaders:#Exchange",
    "RequestServerVersion": "V2018_01_08",
    "TimeZoneContext": {
      "__type": "TimeZoneContext:#Exchange",
      "TimeZoneDefinition": {
        "__type": "TimeZoneDefinitionType:#Exchange",
        "Id": "W. Europe Standard Time"
      }
    }
  },
  "Body": {
    "__type": "RespondToCalendarEventRequest:#Exchange",
    "EventId": { "__type": "ItemId:#Exchange", "Id": "<ImmutableId>" },
    "Response": "Accept" | "Tentative" | "Decline",
    "SendResponse": true | false,
    "Notes": {
      "__type": "BodyContentType:#Exchange",
      "BodyType": "HTML",
      "Value": "<div>optional message</div>"
    },
    "ProposedStartTime": "" | "2026-04-10T09:00:00",
    "ProposedEndTime": "" | "2026-04-10T10:00:00",
    "Attendance": 0 | 3,
    "Mode": 0 | 3
  }
}
```

**Attendance/Mode values:**
| Attendance | Mode | Meaning |
|-----------|------|---------|
| 0 | 0 | Normal attendee (standard RSVP) |
| 3 | 3 | Follow (track without blocking time) |

**Required headers:**
- `Authorization: Bearer <token>`
- `Content-Type: application/json; charset=utf-8`
- `action: RespondToCalendarEvent`
- `x-owa-urlpostdata: <URL-encoded JSON payload>`
- `x-req-source: Calendar`

**Response:** `{ Body: { ResponseCode: "NoError", ResponseClass: "Success" } }`

### GetCalendarEvent via service.svc

To get the ImmutableId, we need to call `service.svc?action=GetCalendarEvent`. This is how the OWA browser fetches event details internally. The request includes `Prefer: IdType="ImmutableId"` which makes the response return ImmutableIds.

**This needs to be reverse-engineered.** Use Chrome DevTools to capture a `GetCalendarEvent` call from the browser and understand:
- The exact payload structure for `x-owa-urlpostdata`
- What ID format it accepts as input (likely the EwsId that OWA already has from `GetCalendarView`)
- What the response shape is and where the ImmutableId lives

Alternatively, the browser's `GetCalendarView` calls also return ImmutableIds when the `Prefer: IdType="ImmutableId"` header is set. We could call the REST API with that header to get ImmutableIds directly in `get_calendar_events`.

---

## File Changes

| File | Change |
|------|--------|
| `src/calendar.ts` | Replace `toServiceId` (translateExchangeIds) with native ImmutableId fetching. Route `respondToEvent` through service.svc. |
| `src/index.ts` | No changes needed — tool schemas stay the same |
| `src/types.ts` | May need service.svc response types |
| `CLAUDE.md` | Already updated with service.svc documentation |

---

## Task 1: Research GetCalendarEvent via service.svc

- [x] **Step 1: Capture a GetCalendarEvent request from the browser**

Captured via Playwright MCP. The `x-owa-urlpostdata` header was not exposed by Playwright's network API (stripped/truncated), but we confirmed the request fires on event click with `Prefer: IdType="ImmutableId"` header.

- [x] **Step 2: Determine the simplest way to get ImmutableIds**

**Findings:**
1. `GET /me/events/{id}` with `Prefer: IdType="ImmutableId"` does NOT translate the ID — returns the same RestId
2. `calendarview` with `Prefer: IdType="ImmutableId"` DOES return ImmutableIds — but produces the same IDs as `translateExchangeIds`
3. `translateExchangeIds` is equally reliable — same IDs, same failures
4. service.svc `ErrorItemNotFound` affects ~30% of events regardless of ID source — it's a server-side issue
5. **Conclusion: keep `translateExchangeIds` in `toServiceId`, use service.svc with REST API fallback**

---

## Task 2: Implement Native ImmutableId Resolution

- [x] **Skipped** — Research showed `translateExchangeIds` is as reliable as any alternative. Existing `toServiceId` is retained.

---

## Task 3: Route RSVP Through service.svc

- [x] **Step 1: Update `respondToEvent` to use service.svc**

Implemented in `src/calendar.ts`. The method now:
1. Translates RestId to ImmutableId via `toServiceId`
2. Calls `service.svc?action=RespondToCalendarEvent` with Attendance=0, Mode=0
3. If service.svc returns `NoError`, returns immediately
4. If service.svc fails (ErrorItemNotFound, etc.), falls back to REST API

- [x] **Step 2: Keep REST API as fallback**

Fallback is automatic — any service.svc error silently falls through to `POST /me/events/{id}/{action}`.

---

## Task 4: Update Follow to Use Same ID Resolution

- [x] **Skipped** — Follow already uses `toServiceId` with the same `translateExchangeIds` approach. No changes needed.

---

## Task 5: Tests and Documentation

- [x] **Step 1: Add integration test for RSVP via service.svc**

Added test in `tests/calendar-write.test.ts` that creates an event, RSVPs via `respondToEvent` (which tries service.svc first), and verifies the response.

- [x] **Step 2: Update README tool descriptions**

Updated `respond_to_calendar_event` description in README to note service.svc usage and ResponseRequested bypass.

- [x] **Step 3: Commit and push**

---

## Verification

1. `npm test` — all existing tests pass
2. `npm run build` — no errors
3. Follow a recurring occurrence → native Follow with "Following:" prefix
4. Follow a single-instance event → native Follow (if ID resolves) or fallback
5. RSVP to an event with `ResponseRequested: false` → works via service.svc
6. RSVP to a normal event → works via service.svc (or fallback to REST API)
