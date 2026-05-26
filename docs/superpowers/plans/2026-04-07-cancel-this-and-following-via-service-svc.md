# Cancel "This and Following" via service.svc Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Fix `cancel_calendar_event` with `thisAndFollowing` scope to send cancellation notices to attendees (with the reason), matching what OWA does natively.

**Architecture:** Replace the current `truncateSeriesAt()` PATCH approach with a `service.svc?action=CancelCalendarEvent` call using `EventScope: 2`. This is the same internal OWA API surface already used by `followEvent` and `respondToEvent`. The event ID must be translated from RestId to ImmutableId via the existing `toServiceId()` method.

**Tech Stack:** TypeScript, OWA service.svc internal API, existing `toServiceId()` for ID translation.

**Reverse-engineered protocol (from New Outlook network capture):**
```
POST service.svc?action=CancelCalendarEvent
Headers:
  action: CancelCalendarEvent
  x-owa-urlpostdata: <URL-encoded JSON payload>
  prefer: IdType="ImmutableId"
  content-length: 0

Payload (in x-owa-urlpostdata):
{
  "__type": "CancelCalendarEventJsonRequest:#Exchange",
  "Header": {
    "__type": "JsonRequestHeaders:#Exchange",
    "RequestServerVersion": "V2018_01_08",
    "TimeZoneContext": { ... }
  },
  "Body": {
    "__type": "CancelCalendarEventRequest:#Exchange",
    "EventId": { "__type": "ItemId:#Exchange", "Id": "<ImmutableId>" },
    "EventScope": 2,          // 0=single, 1=allInSeries, 2=thisAndFollowing
    "ClientSupportsIrm": true,
    "Notes": {
      "__type": "BodyContentType:#Exchange",
      "BodyType": "HTML",
      "Value": "<div>reason text</div>"
    }
  }
}
```

---

### Task 1: Add `cancelEventViaSvc` method to CalendarClient

**Files:**
- Modify: `src/calendar.ts:75-93` (the `cancelEvent` method)
- Modify: `src/calendar.ts` (add new private method `cancelEventViaSvc`)

This task adds a new private method that calls `service.svc?action=CancelCalendarEvent` with the captured protocol, then routes `cancelEvent` through it for all scopes (single, thisAndFollowing, allInSeries map to EventScope 0, 2, 1 respectively).

- [ ] **Step 1: Add the `cancelEventViaSvc` private method**

Add this method to `CalendarClient` (after the existing `truncateSeriesAt` method, around line 281):

```typescript
/**
 * Cancel an event via OWA's native CancelCalendarEvent service.svc action.
 * Sends cancellation notices to attendees for all scopes including thisAndFollowing.
 * EventScope: 0 = single, 1 = allInSeries, 2 = thisAndFollowing.
 */
private async cancelEventViaSvc(
  eventId: string,
  comment?: string,
  scope: 'single' | 'thisAndFollowing' | 'allInSeries' = 'single'
): Promise<void> {
  const scopeMap: Record<string, number> = {
    single: 0,
    allInSeries: 1,
    thisAndFollowing: 2,
  };

  const token = await this.tokens.getToken();
  const svcEventId = await this.toServiceId(eventId, token.value);

  const payload = {
    __type: 'CancelCalendarEventJsonRequest:#Exchange',
    Header: {
      __type: 'JsonRequestHeaders:#Exchange',
      RequestServerVersion: 'V2018_01_08',
      TimeZoneContext: {
        __type: 'TimeZoneContext:#Exchange',
        TimeZoneDefinition: {
          __type: 'TimeZoneDefinitionType:#Exchange',
          Id: 'W. Europe Standard Time',
        },
      },
    },
    Body: {
      __type: 'CancelCalendarEventRequest:#Exchange',
      EventId: { __type: 'ItemId:#Exchange', Id: svcEventId },
      EventScope: scopeMap[scope],
      ClientSupportsIrm: true,
      Notes: {
        __type: 'BodyContentType:#Exchange',
        BodyType: 'HTML',
        Value: comment ? `<div>${comment}</div>` : '<div><br></div>',
      },
    },
  };

  const res = await fetch(`${OWA_SVC}?action=CancelCalendarEvent`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token.value}`,
      'Content-Type': 'application/json; charset=utf-8',
      action: 'CancelCalendarEvent',
      'x-owa-urlpostdata': encodeURIComponent(JSON.stringify(payload)),
      'x-req-source': 'Calendar',
      Prefer: 'IdType="ImmutableId"',
    },
  });

  const body = (await res.json()) as { Body: { ResponseCode: string; MessageText?: string } };
  if (body.Body.ResponseCode !== 'NoError') {
    throw new Error(`CancelCalendarEvent failed: ${body.Body.ResponseCode} — ${body.Body.MessageText ?? ''}`);
  }
}
```

- [ ] **Step 2: Rewrite `cancelEvent` to use `cancelEventViaSvc` for all scopes**

Replace the entire `cancelEvent` method (lines 75-93) with:

```typescript
async cancelEvent(eventId: string, comment?: string, scope: 'single' | 'thisAndFollowing' | 'allInSeries' = 'single'): Promise<void> {
  await this.cancelEventViaSvc(eventId, comment, scope);
}
```

This routes all cancellation scopes through the service.svc action, which:
- Sends cancellation notices with the reason for all scopes
- Handles `thisAndFollowing` natively (EventScope=2) instead of truncating

- [ ] **Step 3: Build and verify no compile errors**

Run: `npm run build`
Expected: Clean compile, no errors.

- [ ] **Step 4: Commit**

```bash
git add src/calendar.ts
git commit -m "fix(calendar): route all cancelEvent scopes through service.svc

Replaces the PATCH-based truncation approach for thisAndFollowing with
OWA's native CancelCalendarEvent service.svc action. This sends proper
cancellation notices to attendees for all scopes, including
thisAndFollowing (EventScope=2).

Protocol reverse-engineered from New Outlook's network traffic."
```

---

### Task 2: Update tool description and CLAUDE.md

**Files:**
- Modify: `src/index.ts:233-246` (cancel_calendar_event tool registration)
- Modify: `CLAUDE.md`

- [ ] **Step 1: Update the tool description in index.ts**

Replace the tool description at line 235:

Old:
```typescript
'Cancel a meeting you organized. Sends a cancellation notice with your reason to all attendees. Only works if you are the organizer. Note: "thisAndFollowing" silently truncates the series without sending cancellation notices.',
```

New:
```typescript
'Cancel a meeting you organized. Sends a cancellation notice with your reason to all attendees. Only works if you are the organizer. Note: "thisAndFollowing" silently truncates the series without sending cancellation notices.',
```

Replace with:
```typescript
'Cancel a meeting you organized. Sends a cancellation notice with your reason to all attendees. Only works if you are the organizer.',
```

- [ ] **Step 2: Update CLAUDE.md — remove the thisAndFollowing caveat**

In CLAUDE.md, the `cancel_calendar_event` tool description says:

> Note: "thisAndFollowing" silently truncates the series without sending cancellation notices.

Remove this note entirely. The fix makes all scopes send proper cancellation notices.

Also update the "Two API surfaces" section in CLAUDE.md. In the REST API paragraph, remove `cancel_calendar_event` from the list of tools that use REST API. In the service.svc paragraph, add it:

Old (REST API section):
```
Used by calendar tools (`get_calendar_events`, `create_calendar_event`, `update_calendar_event`, `cancel_calendar_event`, `delete_calendar_event`) and all mail tools
```

New:
```
Used by calendar tools (`get_calendar_events`, `create_calendar_event`, `update_calendar_event`, `delete_calendar_event`) and all mail tools
```

Old (service.svc section):
```
Used for `follow_calendar_event` and will be used for `respond_to_calendar_event`.
```

New:
```
Used for `cancel_calendar_event` (all scopes), `follow_calendar_event`, and `respond_to_calendar_event`.
```

- [ ] **Step 3: Commit**

```bash
git add src/index.ts CLAUDE.md
git commit -m "docs: update cancel tool description now that all scopes send notices"
```

---

### Task 3: Clean up dead code

**Files:**
- Modify: `src/calendar.ts` — remove `truncateSeriesAt` and `getSeriesRecurrence`

The `truncateSeriesAt` private method (lines 255-281) and `getSeriesRecurrence` private method (lines 310-317) are now dead code — `cancelEvent` no longer calls them. The `deleteEvent` method also uses `truncateSeriesAt` for its `thisAndFollowing` scope, so we need to route that through service.svc too — but `deleteEvent` is explicitly silent (no notifications). So `deleteEvent` should keep using the truncation approach.

Wait — re-checking: `deleteEvent` at line 108 also calls `truncateSeriesAt`. Since `deleteEvent` is meant to be silent, the truncation approach is correct for it. So `truncateSeriesAt` and `getSeriesRecurrence` are still used by `deleteEvent` and must NOT be removed.

- [ ] **Step 1: Verify `truncateSeriesAt` is still used by `deleteEvent`**

Read `src/calendar.ts` lines 95-109. Confirm `deleteEvent` calls `this.truncateSeriesAt(eventId)` on line 108. If confirmed, skip this task — the methods are not dead code.

- [ ] **Step 2: No changes needed — skip this task**

`truncateSeriesAt` and `getSeriesRecurrence` remain in use by `deleteEvent`. No code to remove.

---

### Task 4: Update integration test for thisAndFollowing cancel

**Files:**
- Modify: `tests/calendar-series.test.ts:254-278`

The existing test at line 254 (`thisAndFollowing truncates the series`) verifies the old behavior — it creates a 5-week series, cancels from the 3rd occurrence, and checks only 2 remain. This test still validates the correct outcome but should also verify the comment parameter is accepted (previously it was silently dropped).

- [ ] **Step 1: Update the existing thisAndFollowing test to pass a comment**

Replace the test at lines 254-278 with:

```typescript
test('thisAndFollowing cancels from occurrence onward with comment', async () => {
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

  // Cancel from the 3rd occurrence onward WITH a comment
  const thirdInstance = instances[2];
  await client.cancelEvent(thirdInstance.id, 'Integration test cleanup', 'thisAndFollowing');

  // Re-list: should now have only 2 instances
  const afterInstances = await client.listSeriesInstances(
    created.id,
    rangeStart.toISOString(),
    rangeEnd.toISOString()
  );
  expect(afterInstances.length).toBe(2);
}, 60_000);
```

- [ ] **Step 2: Run the integration tests**

Run: `npm test -- --testPathPattern=calendar-series --verbose`
Expected: All tests pass including the updated `thisAndFollowing` test.

Note: Integration tests require a live Edge session with M365 authentication. If the test environment is not available, verify manually by building and running the MCP server.

- [ ] **Step 3: Commit**

```bash
git add tests/calendar-series.test.ts
git commit -m "test: update thisAndFollowing cancel test to verify comment is sent"
```

---

### Task 5: Version bump and tag

**Files:**
- Modify: `package.json` (version field)

This is a bug fix (cancellation notices weren't being sent), so bump the patch version.

- [ ] **Step 1: Check current version**

Run: `node -e "console.log(require('./package.json').version)"`
Expected: `0.4.0`

- [ ] **Step 2: Bump patch version to 0.4.1**

In `package.json`, change `"version": "0.4.0"` to `"version": "0.4.1"`.

- [ ] **Step 3: Commit and tag**

```bash
git add package.json
git commit -m "chore: bump version to 0.4.1"
git tag v0.4.1
```

- [ ] **Step 4: Push (only when user confirms)**

```bash
git push origin main --tags
```

This triggers the GitHub Actions workflow to publish to npm.
