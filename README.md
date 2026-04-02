# owa-mcp

A [Model Context Protocol (MCP)](https://modelcontextprotocol.io) server that gives Claude Code read access to your Microsoft Outlook calendar — **without requiring an Azure app registration**.

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

Restart Claude Code. You should now have a `get_calendar_events` tool available.

## Available Tools

### `get_calendar_events`

Returns calendar events in a time range.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `startDateTime` | string | yes | ISO 8601 start, e.g. `2026-04-07T00:00:00Z` |
| `endDateTime` | string | yes | ISO 8601 end, e.g. `2026-04-14T00:00:00Z` |
| `maxResults` | number | no | Max events to return (default 50, max 100) |
| `timezone` | string | no | IANA timezone, e.g. `Europe/Berlin` (default UTC) |

Example prompt: *"What meetings do I have next week?"*

## Troubleshooting

**Token acquisition times out**
Open Edge, navigate to outlook.office.com, confirm you can see your calendar. The session may have expired — sign in again.

**`ErrorAccessDenied` on calendar API**
The intercepted token didn't carry calendar scope. This is rare; try quitting all Edge windows and restarting.

**Headless browser opens a visible window**
This shouldn't happen normally. If it does, check that no other Playwright process is holding the Edge profile directory lock.

## Roadmap

- [ ] `create_calendar_event`
- [ ] `update_calendar_event`
- [ ] `delete_calendar_event`
- [ ] `get_emails` (Mail.ReadWrite scope is already present in the token)
- [ ] `send_email`

## License

Apache 2.0 — see [LICENSE](LICENSE).
