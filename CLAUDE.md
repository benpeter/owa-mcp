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

The OWA REST API base is `https://outlook.office.com/api/v2.0`. It mirrors the Microsoft Graph API shape closely — most Graph calendar/mail docs apply with `Subject`/`Start`/`End` casing instead of `subject`/`start`/`end`.

## Known limitations

- macOS only (Edge profile path is hardcoded to Mac location)
- Requires Edge installed at `/Applications/Microsoft Edge.app`
- Token acquisition takes ~8–10 seconds on cold start (headless browser launch)
- If the Edge session expires (usually after weeks of inactivity), the user must sign in to outlook.office.com in Edge again
