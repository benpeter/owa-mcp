# owa-mcp

A [Model Context Protocol (MCP)](https://modelcontextprotocol.io) server that gives Claude Code full access to your Microsoft Outlook calendar and email — **without requiring an Azure app registration**.

## How it works

The server drives a dedicated, persistent Google Chrome browser via [chrome-devtools-mcp](https://github.com/ChromeDevTools/chrome-devtools-mcp), using its own Chrome profile (kept separate from your daily-driver Chrome) so the login session survives across restarts. It navigates that browser to Outlook Web (outlook.office.com) and watches the network requests it makes. When Outlook Web loads, it issues Bearer tokens for its own internal API calls; this server intercepts those tokens and reuses them against the `outlook.office.com/api/v2.0` REST endpoint.

The result: full `Calendars.ReadWrite` and `Mail.ReadWrite` scope with no OAuth app registration, no client ID, and no IT involvement — as long as you can sign in to Microsoft 365 in the automation browser.

**Tokens expire after ~80 minutes.** Rather than relaunching a browser for every refresh, the server keeps a single chrome-devtools-mcp connection (and its browser) alive across acquisitions and simply re-navigates to Outlook Web to trigger a fresh token request. The Chrome window is visible (never headless) from the moment the browser first starts, so it's there for you to sign in whenever needed. If no active session is found within 60 seconds, the request fails with a message asking you to sign in and retry — because the connection stays alive, you can take as long as you need (MFA included) and just retry once done, without losing the browser session.

## Why this approach

Many enterprise Microsoft 365 tenants enforce Conditional Access policies that block third-party OAuth flows (e.g., Azure CLI, custom app registrations) on managed devices. The browser-session interception approach works because it piggybacks on an authentication flow that already satisfies those policy requirements — the same one used by Outlook Web itself. This has been confirmed working with Chrome SSO in at least one Conditional-Access-enforced corp environment; whether it works in yours depends on your organization's specific policies.

## Prerequisites

- macOS (tested on macOS 15)
- [Google Chrome](https://www.google.com/chrome/) installed
- [chrome-devtools-mcp](https://github.com/ChromeDevTools/chrome-devtools-mcp) available via `npx` — verify with `npx chrome-devtools-mcp@latest --version`; the server spawns it automatically on first use, no manual setup needed
- Signed in to Microsoft 365 (the first token request opens a visible Chrome window — sign in there if prompted)
- Node.js 20+

## Installation

```bash
claude mcp add owa -s user -- npx owa-mcp
```

That's it. Restart Claude Code — you should see calendar tools available.

<details>
<summary>Manual installation (alternative)</summary>

```bash
git clone https://github.com/benpeter/owa-mcp
cd owa-mcp
npm install
npm run build
claude mcp add owa -s user -- node /absolute/path/to/owa-mcp/dist/index.js
```

</details>

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
| `hideAttendees` | boolean | no | Hide attendee list from other attendees (default false) |
| `responseRequested` | boolean | no | Request RSVPs from attendees (default true) |
| `reminderMinutes` | number | no | Reminder minutes before start. Omit for Outlook default, 0 to disable |
| `recurrence` | object | no | Make this a recurring event. See [Recurrence](#recurrence-object) below |

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
| `isPrivate` | boolean | no | Mark as private |
| `hideAttendees` | boolean | no | Hide attendee list from other attendees |
| `responseRequested` | boolean | no | Request RSVPs from attendees |
| `reminderMinutes` | number | no | Reminder minutes before start. 0 to disable |
| `recurrence` | object | no | Change the recurrence pattern. Only applies to series master events. See [Recurrence](#recurrence-object) below |

### `cancel_calendar_event`

Cancel a meeting you organized. Sends cancellation with reason to attendees. Supports recurring series operations.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `eventId` | string | yes | Event ID |
| `reason` | string | no | Cancellation reason sent to attendees |
| `scope` | string | no | `single` (default), `thisAndFollowing`, or `allInSeries` |

### `delete_calendar_event`

Remove an event from your calendar silently (no notification sent). Supports recurring series operations.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `eventId` | string | yes | Event ID |
| `scope` | string | no | `single` (default), `thisAndFollowing`, or `allInSeries` |

### `respond_to_calendar_event`

RSVP to a meeting: accept, tentatively accept, or decline. Uses OWA's internal service.svc API when possible, which works even when the organizer has disabled response requests (`ResponseRequested: false`). Falls back to the standard REST API if the internal API can't resolve the event.

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
| `comment` | string | no | Optional message included in the follow notification to the organizer |
| `timezone` | string | no | Timezone for returned event |

### `get_series_master`

Inspect the master event of a recurring series. Returns recurrence pattern, cancelled occurrences, and full event details. Accepts any event ID from the series.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `eventId` | string | yes | Any event ID from the series (resolved automatically) |
| `timezone` | string | no | IANA timezone (default UTC) |

### `list_series_instances`

List all occurrences of a recurring series within a date range. Accepts any event ID from the series.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `eventId` | string | yes | Any event ID from the series (resolved automatically) |
| `startDateTime` | string | yes | ISO 8601 start |
| `endDateTime` | string | yes | ISO 8601 end |
| `timezone` | string | no | IANA timezone (default UTC) |

### `list_mail_folders`

List all mail folders in the mailbox, or child folders of a specific folder.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `parentFolderId` | string | no | List children of this folder. If omitted, lists top-level folders |

### `get_emails`

Get emails from a specific mailbox folder with optional filtering.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `folderId` | string | no | Folder ID or well-known name (Inbox, Drafts, SentItems, DeletedItems). Default: Inbox |
| `filter` | string | no | `all`, `unread`, `flagged`, `today`, `this_week` |
| `limit` | number | no | Max results (default 20, max 500) |
| `pageToken` | string | no | Pagination token from previous response |

### `search_emails`

Search emails using full-text query OR structured filters (mutually exclusive).

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `query` | string | no | Full-text search query. Cannot combine with structured filters |
| `from` | string | no | Filter by sender email |
| `subject` | string | no | Filter by subject (contains) |
| `receivedAfter` | string | no | ISO 8601 datetime |
| `receivedBefore` | string | no | ISO 8601 datetime |
| `folderId` | string | no | Scope search to folder |
| `limit` | number | no | Max results (default 20, max 500) |

### `get_email`

Read a single email with full body content and attachment metadata.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `messageId` | string | yes | Message ID |
| `format` | string | no | `text` (default) or `html` |

### `get_attachment`

Download an email attachment to disk.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `messageId` | string | yes | Message ID |
| `attachmentId` | string | yes | Attachment ID from get_email response |

### `send_email`

Compose and send a new email in one step. For more control, use `create_draft` + `update_draft` + `send_draft`.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `to` | array | yes | `[{ email, name? }]` |
| `subject` | string | yes | Subject line |
| `body` | string | yes | Body content |
| `bodyType` | string | no | `text` (default) or `html` |
| `cc` | array | no | CC recipients `[{ email, name? }]` |
| `bcc` | array | no | BCC recipients |
| `importance` | string | no | Low, Normal (default), High |
| `saveToSentItems` | boolean | no | Save to Sent Items (default true) |

### `create_draft`

Create a new email draft saved to Drafts folder.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `to` | array | yes | `[{ email, name? }]` |
| `subject` | string | yes | Subject line |
| `body` | string | yes | Body content |
| `bodyType` | string | no | `text` (default) or `html` |
| `cc` | array | no | CC recipients |
| `bcc` | array | no | BCC recipients |
| `importance` | string | no | Low, Normal (default), High |

### `create_reply_draft`

Create a draft reply to the sender. Returns draft with pre-filled recipients, quoted body, "RE:" subject.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `messageId` | string | yes | Message ID |

### `create_reply_all_draft`

Create a draft reply-all. Returns draft with all original recipients, quoted body, "RE:" subject.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `messageId` | string | yes | Message ID |

### `create_forward_draft`

Create a draft forward. Returns draft with quoted body, "FW:" subject, no To recipients.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `messageId` | string | yes | Message ID |

### `update_draft`

Modify a draft before sending. Can change subject, body, recipients, importance.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `messageId` | string | yes | Draft message ID |
| `subject` | string | no | New subject |
| `body` | string | no | New body content |
| `bodyType` | string | no | `text` or `html` |
| `toRecipients` | array | no | Replace all To recipients |
| `ccRecipients` | array | no | Replace all CC recipients |
| `bccRecipients` | array | no | Replace all BCC recipients |
| `importance` | string | no | Low, Normal, High |

### `send_draft`

Send a draft message. Moves from Drafts to Sent Items.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `messageId` | string | yes | Draft message ID |

### `move_email`

Move a message to a different folder. Returns the moved message (with updated ID).

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `messageId` | string | yes | Message ID |
| `destinationId` | string | yes | Folder ID or well-known name (Inbox, Drafts, SentItems, DeletedItems, Archive) |

### `delete_email`

Delete a message (moves to Deleted Items).

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `messageId` | string | yes | Message ID |

### `update_email`

Update email properties: mark as read/unread, flag/unflag.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `messageId` | string | yes | Message ID |
| `isRead` | boolean | no | Set read (true) or unread (false) |
| `flagStatus` | string | no | `NotFlagged`, `Flagged`, or `Complete` |

Example prompts:
- *"What meetings do I have next week?"*
- *"Create a 30-minute meeting with Jane tomorrow at 2pm"*
- *"Decline the ECCN sync with a note that I'm on vacation"*
- *"Follow the Analytics Tech Call so I can see it on my calendar"*
- *"Cancel all future occurrences of the weekly sync starting from next week"*
- *"What's the recurrence pattern for the Monday standup?"*
- *"Create a weekly team sync every Tuesday at 10am for the next 3 months"*
- *"Show me unread emails from today"*
- *"Reply to that email from Sarah and add Bob to CC"*
- *"Forward the Q3 report to the finance team"*
- *"Mark all emails from the newsletter as read"*

## Recurrence Object

Used by `create_calendar_event` and `update_calendar_event` to define recurring events.

```json
{
  "recurrence": {
    "pattern": {
      "type": "weekly",
      "interval": 1,
      "daysOfWeek": ["Monday", "Wednesday", "Friday"]
    },
    "range": {
      "type": "endDate",
      "startDate": "2026-04-07",
      "endDate": "2026-07-07"
    }
  }
}
```

**Pattern types:** `daily`, `weekly`, `absoluteMonthly`, `relativeMonthly`, `absoluteYearly`, `relativeYearly`

| Pattern Field | Type | Description |
|---------------|------|-------------|
| `type` | string | Required. Pattern type |
| `interval` | number | Required. Interval between occurrences (1 = every, 2 = every other) |
| `daysOfWeek` | string[] | Days for weekly/relative patterns |
| `dayOfMonth` | number | Day of month for absoluteMonthly/absoluteYearly |
| `month` | number | Month (1-12) for yearly patterns |
| `index` | string | Week index for relative patterns: first, second, third, fourth, last |
| `firstDayOfWeek` | string | First day of week (default Sunday) |

**Range types:** `endDate`, `numbered`, `noEnd`

| Range Field | Type | Description |
|-------------|------|-------------|
| `type` | string | Required. How the series ends |
| `startDate` | string | Required. Series start (YYYY-MM-DD) |
| `endDate` | string | End date (required for `endDate` type) |
| `numberOfOccurrences` | number | Count (required for `numbered` type) |
| `recurrenceTimeZone` | string | Timezone for recurrence dates |

## Troubleshooting

**Token acquisition times out / no active session found**
The automation Chrome window is already open (it becomes visible as soon as the browser starts, not just on timeout) and has navigated to Outlook Web. Switch to it and sign in there (MFA included) at your own pace, then simply retry the request — the browser connection is kept alive, so the same session is reused and you won't need to sign in again on subsequent calls.

**`ErrorAccessDenied` on calendar API**
The intercepted token didn't carry calendar scope. This is rare; try quitting any Chrome windows using the automation profile (`~/Library/Application Support/owa-mcp/chrome-profile`) and retrying — the server will relaunch it automatically.

## License

Apache 2.0 — see [LICENSE](LICENSE).
