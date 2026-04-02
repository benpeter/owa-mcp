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

- [ ] **Step 1: Capture a GetCalendarEvent request from the browser**

Use Chrome DevTools to capture a `service.svc?action=GetCalendarEvent` request. Decode the `x-owa-urlpostdata` header to understand:
- The payload structure
- What input ID format is used
- What the response contains (specifically the ImmutableId field name)

- [ ] **Step 2: Determine the simplest way to get ImmutableIds**

Options:
1. Call `service.svc?action=GetCalendarEvent` with the REST API's EwsId
2. Call the REST API's `GET /me/events/{id}` with `Prefer: IdType="ImmutableId"` header (this is what the REST API uses for response formatting — may also work for ID format)
3. Use a separate `service.svc?action=GetCalendarView` to get events with ImmutableIds from the start

Option 2 is the simplest if it works. Test it.

---

## Task 2: Implement Native ImmutableId Resolution

- [ ] **Step 1: Replace `toServiceId` method**

Replace the current `toServiceId` (which uses `translateExchangeIds`) with a method that gets the ImmutableId reliably. The implementation depends on Task 1 findings.

If Option 2 works:
```typescript
private async toServiceId(restId: string, token: string): Promise<string> {
  // Fetch event with Prefer: IdType="ImmutableId" to get the native ImmutableId
  const res = await fetch(`${OWA_BASE}/me/events/${restId}?$select=Id`, {
    headers: {
      Authorization: `Bearer ${token}`,
      Prefer: 'IdType="ImmutableId"',
    },
  });
  const data = await res.json();
  return data.Id; // Already in ImmutableId format
}
```

- [ ] **Step 2: Verify ImmutableId works for Follow on problematic events**

Test with events that previously failed (single-instance events with `-` in REST ID).

---

## Task 3: Route RSVP Through service.svc

- [ ] **Step 1: Update `respondToEvent` to use service.svc**

```typescript
async respondToEvent(
  eventId: string,
  action: RsvpAction,
  payload: OwaRsvpPayload = {}
): Promise<void> {
  const token = await this.tokens.getToken();
  const svcEventId = await this.toServiceId(eventId, token.value);

  const responseMap: Record<RsvpAction, string> = {
    accept: 'Accept',
    tentativelyaccept: 'Tentative',
    decline: 'Decline',
  };

  const svcPayload = {
    __type: 'RespondToCalendarEventJsonRequest:#Exchange',
    Header: { /* standard header */ },
    Body: {
      __type: 'RespondToCalendarEventRequest:#Exchange',
      EventId: { __type: 'ItemId:#Exchange', Id: svcEventId },
      Response: responseMap[action],
      SendResponse: payload.SendResponse ?? true,
      Notes: payload.Comment
        ? { __type: 'BodyContentType:#Exchange', BodyType: 'HTML', Value: `<div>${payload.Comment}</div>` }
        : undefined,
      ProposedStartTime: payload.ProposedNewTime?.Start?.DateTime ?? '',
      ProposedEndTime: payload.ProposedNewTime?.End?.DateTime ?? '',
      Attendance: 0,
      Mode: 0,
    },
  };

  // Call service.svc
  // Fall back to REST API if service.svc fails
}
```

- [ ] **Step 2: Keep REST API as fallback**

If `service.svc` fails (e.g., ID resolution issue), fall back to the current REST API `POST /me/events/{id}/{action}` approach. This ensures the tool always works, even if degraded (REST API will still reject `ResponseRequested: false` events).

---

## Task 4: Update Follow to Use Same ID Resolution

- [ ] **Step 1: Update `followEvent` to use new `toServiceId`**

Replace the current `translateExchangeIds`-based `toServiceId` with the new implementation from Task 2.

- [ ] **Step 2: Test Follow on previously-failing single-instance events**

---

## Task 5: Tests and Documentation

- [ ] **Step 1: Add integration test for RSVP via service.svc**

Test accepting an event where `ResponseRequested: false` (previously failed).

- [ ] **Step 2: Update README tool descriptions**

Update `respond_to_calendar_event` description to note it works even when organizer disabled responses.

- [ ] **Step 3: Commit and push**

---

## Verification

1. `npm test` — all existing tests pass
2. `npm run build` — no errors
3. Follow a recurring occurrence → native Follow with "Following:" prefix
4. Follow a single-instance event → native Follow (if ID resolves) or fallback
5. RSVP to an event with `ResponseRequested: false` → works via service.svc
6. RSVP to a normal event → works via service.svc (or fallback to REST API)
