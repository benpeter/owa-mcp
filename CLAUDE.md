# CLAUDE.md — owa-mcp

## What this project is

An MCP server that exposes Microsoft Outlook calendar data to Claude Code. It works by intercepting Bearer tokens that Outlook Web (outlook.office.com) emits when loaded in a headless Playwright-controlled Microsoft Edge browser using the user's existing Edge profile.

## Why this architecture exists

The standard approach — registering an Azure app and using the Microsoft Graph API — requires IT admin consent in enterprise Microsoft 365 tenants with Conditional Access policies. Azure CLI authentication is also blocked on managed devices by those same policies.

The solution: Edge is already signed in to M365 and satisfies all Conditional Access requirements. Playwright launches Edge headlessly using the real profile directory (`~/Library/Application Support/Microsoft Edge`). Outlook Web makes API calls using Bearer tokens issued by Microsoft's own OWA app (`appid: 9199bf20-a13f-4107-85dc-02114787ef48`). Those tokens carry extensive delegated scopes including `Calendars.ReadWrite`, `Mail.ReadWrite`, `Files.ReadWrite.All`, and more. The server intercepts them via `page.on('request')` and reuses them against `https://outlook.office.com/api/v2.0`.

## Token details

- Intercepted from requests to `outlook.office.com/owa/service.svc`
- Lifetime: ~80 minutes (tracked via JWT `exp` claim)
- Auto-refreshed 5 minutes before expiry
- Scope confirmed in production: `Calendars.ReadWrite`, `Mail.ReadWrite`, `Contacts.ReadWrite`, `Files.ReadWrite.All`, `Chat.Read`, and ~60 more

## Key files

| File | Purpose |
|------|---------|
| `src/auth.ts` | `TokenManager` — launches headless Edge, intercepts token, caches with expiry |
| `src/calendar.ts` | `CalendarClient` — calls `outlook.office.com/api/v2.0/me/calendarview` |
| `src/types.ts` | Shared types: `OwaToken`, `CalendarEvent`, `OwaCalendarViewResponse` |
| `src/index.ts` | MCP server, tool registration via `@modelcontextprotocol/sdk` |

## Development

```bash
npm install
npm run dev          # run with tsx (no build needed)
npm run build        # compile to dist/
npm test             # integration tests (require Edge + M365 session)
```

Tests are integration tests — they need a live Edge session. There are no unit tests with mocks because the auth flow is inherently side-effectful.

## Adding new tools

1. Add any new API methods to `src/calendar.ts` (or create `src/mail.ts`, `src/contacts.ts`, etc.)
2. Register the tool in `src/index.ts` using `server.tool(name, description, zodSchema, handler)`
3. Add an integration test in `tests/`

**Tool output convention:** All tools must return structured JSON, not formatted text. Return the data array (or object) directly via `JSON.stringify(data, null, 2)`. The consuming LLM is better at reasoning over structured data and can format it however suits the user's question. Never pre-format, truncate fields, or drop metadata — pass through the full normalized object.

The OWA REST API base is `https://outlook.office.com/api/v2.0`. It mirrors the Microsoft Graph API shape closely — most Graph calendar/mail docs apply with `Subject`/`Start`/`End` casing instead of `subject`/`start`/`end`.

## Two API surfaces

This project uses **two different OWA API surfaces**:

### 1. REST API (`outlook.office.com/api/v2.0`)
Standard REST endpoints for CRUD operations. Used by `get_calendar_events`, `create_calendar_event`, `update_calendar_event`, `cancel_calendar_event`, `delete_calendar_event`. Event IDs are in EwsId/RestId format (base64url, start with `AAMkA`).

### 2. OWA service.svc (`outlook.office.com/owa/service.svc`)
Internal OWA endpoint used by the Outlook Web client. Payload goes in the `x-owa-urlpostdata` header (URL-encoded JSON), not the request body (content-length=0). Used for `follow_calendar_event` and will be used for `respond_to_calendar_event`.

**Key differences from REST API:**
- Bypasses `ResponseRequested: false` — can RSVP to events where the REST API returns "organizer hasn't requested a response"
- Uses `RestImmutableEntryId` format for event IDs (base64, start with `AAkA`), NOT the REST API's `RestId` format
- Translation via `POST /api/beta/me/translateExchangeIds` with `SourceIdType: 'RestId'`, `TargetIdType: 'RestImmutableEntryId'` — then convert base64url to standard base64 (`-` → `+`, `_` → `/`)
- **ID translation is unreliable for some events** (already-followed events, some single-instance events return `ErrorItemNotFound`). The correct approach is to fetch ImmutableIds directly via `service.svc?action=GetCalendarEvent` with `Prefer: IdType="ImmutableId"` header, which is what the OWA browser client does internally.
- Supports `Attendance` and `Mode` fields not available in the REST API:
  - `Attendance: 0, Mode: 0` = normal attendee (Accept/Tentative/Decline)
  - `Attendance: 3, Mode: 3` = Follow (track without RSVPing)

**Follow protocol (reverse-engineered from New Outlook):**
```
POST service.svc?action=RespondToCalendarEvent
x-owa-urlpostdata: { Body: { Response: "Tentative", Attendance: 3, Mode: 3, SendResponse: true/false, Notes: { Value: "<div>message</div>" } } }
```
- Recurring occurrences: `SendResponse: true` (organizer gets "is following" notification, subject prefixed "Following:")
- Single-instance events: `SendResponse: false` (no notification, but subject still prefixed "Following:")

## Known limitations

- macOS only (Edge profile path is hardcoded to Mac location)
- Requires Edge installed at `/Applications/Microsoft Edge.app`
- Token acquisition takes ~8–10 seconds on cold start (headless browser launch)
- If the Edge session expires (usually after weeks of inactivity), the user must sign in to outlook.office.com in Edge again
