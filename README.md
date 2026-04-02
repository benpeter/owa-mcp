# owa-mcp

A [Model Context Protocol (MCP)](https://modelcontextprotocol.io) server that gives Claude Code full access to your Microsoft Outlook calendar — **without requiring an Azure app registration**.

## How it works

Microsoft Outlook Web (outlook.office.com) runs inside a Playwright-controlled headless Microsoft Edge browser that uses your existing, signed-in Edge profile. When Outlook Web loads, it issues Bearer tokens for its own internal API calls. This server intercepts those tokens and reuses them against the `outlook.office.com/api/v2.0` REST endpoint.

The result: full `Calendars.ReadWrite` scope with no OAuth app registration, no client ID, and no IT involvement — as long as you are already signed in to Microsoft 365 in your Edge browser.

**Tokens expire after ~80 minutes.** The server refreshes automatically by re-launching the headless browser in the background.

## Why this approach

Many enterprise Microsoft 365 tenants enforce Conditional Access policies that block third-party OAuth flows (e.g., Azure CLI, custom app registrations). Managed devices may restrict which apps can authenticate. The browser-session interception approach works because it piggybacks on an authentication flow that already satisfies all policy requirements — the same one used by Outlook Web itself.

## Prerequisites

- macOS (tested on macOS 15)
- [Microsoft Edge](https://www.microsoft.com/en-us/edge) installed at `/Applications/Microsoft Edge.app`
- Signed in to Microsoft 365 in Edge (open Edge, go to outlook.office.com, confirm you see your calendar)
- Node.js 20+

## Installation

```bash
git clone https://github.com/benpeter/owa-mcp
cd owa-mcp
npm install
npm run build
```

## Claude Code Configuration

Add to `~/.claude/settings.json`:

```json
{
  "mcpServers": {
    "owa": {
      "command": "node",
      "args": ["/absolute/path/to/owa-mcp/dist/index.js"]
    }
  }
}
```

Restart Claude Code. You should now have calendar tools available.

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
| `timezone` | string | no | Timezone for returned event |

Example prompts:
- *"What meetings do I have next week?"*
- *"Create a 30-minute meeting with Jane tomorrow at 2pm"*
- *"Decline the ECCN sync with a note that I'm on vacation"*
- *"Follow the Analytics Tech Call so I can see it on my calendar"*

## Troubleshooting

**Token acquisition times out**
Open Edge, navigate to outlook.office.com, confirm you can see your calendar. The session may have expired — sign in again.

**`ErrorAccessDenied` on calendar API**
The intercepted token didn't carry calendar scope. This is rare; try quitting all Edge windows and restarting.

**Headless browser opens a visible window**
This shouldn't happen normally. If it does, check that no other Playwright process is holding the Edge profile directory lock.

## Roadmap

- [x] `create_calendar_event`
- [x] `update_calendar_event`
- [x] `cancel_calendar_event` (with reason)
- [x] `respond_to_calendar_event` (accept / tentative / decline)
- [x] `follow_calendar_event`
- [x] `delete_calendar_event`
- [ ] `get_emails` (Mail.ReadWrite scope is already present in the token)
- [ ] `send_email`

## License

Apache 2.0 — see [LICENSE](LICENSE).
