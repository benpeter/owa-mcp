# CLAUDE.md — owa-mcp

## What this project is

An MCP server that exposes Microsoft Outlook calendar and email data to Claude Code. It works by intercepting Bearer tokens that Outlook Web (outlook.office.com) emits when loaded in a Chrome browser automated via [chrome-devtools-mcp](https://github.com/ChromeDevTools/chrome-devtools-mcp), using a dedicated persistent Chrome profile.

## Why this architecture exists

The standard approach — registering an Azure app and using the Microsoft Graph API — requires IT admin consent in enterprise Microsoft 365 tenants with Conditional Access policies. Azure CLI authentication is also blocked on managed devices by those same policies.

The solution: a real Chrome browser signed in to M365 satisfies all Conditional Access requirements. `src/auth.ts` drives Chrome as an MCP client — it spawns `npx chrome-devtools-mcp@latest` via `@modelcontextprotocol/sdk`'s `Client` + `StdioClientTransport` and calls its tools (`navigate_page`, `list_network_requests`, `get_network_request`) to observe the browser. chrome-devtools-mcp is a hard prerequisite: there is no fallback, and if it fails to start the server throws an error telling the user to verify it with `npx chrome-devtools-mcp@latest --version` and confirm Chrome is installed. The browser runs against a dedicated, persistent Chrome profile at `~/Library/Application Support/owa-mcp/chrome-profile`, kept separate from the user's daily-driver Chrome and from any chrome-devtools-mcp instance Claude Code itself might have open — this avoids user-data-dir lock conflicts. It is not headless, so a visible window is available whenever the user needs to sign in, and it is not `--isolated`, so login cookies persist across restarts. Outlook Web makes API calls using Bearer tokens issued by Microsoft's own OWA app (`appid: 9199bf20-a13f-4107-85dc-02114787ef48`). Those tokens carry extensive delegated scopes including `Calendars.ReadWrite`, `Mail.ReadWrite`, `Files.ReadWrite.All`, and more. The server extracts them from network requests captured via chrome-devtools-mcp and reuses them against `https://outlook.office.com/api/v2.0`.

Microsoft Edge is no longer used anywhere in this project.

## Token details

- Intercepted from requests whose path starts with `/owa/service.svc`, on any hostname in an explicit allowlist (`outlook.office.com`, `outlook.office365.com`, `outlook.cloud.microsoft`) — checked via `new URL()` with an exact hostname match, not a substring, since a page reached during interactive sign-in could otherwise spoof the path on an attacker-controlled domain. The hostname isn't hardcoded to one value because `outlook.office.com/calendar/view/...` 302-redirects to `outlook.cloud.microsoft` for some tenants (Microsoft's ongoing domain consolidation) — if Microsoft adds another domain, extend `OWA_TOKEN_HOSTS` in `src/auth.ts`, don't loosen the match.
- Lifetime: ~80 minutes (tracked via JWT `exp` claim)
- Auto-refreshed 5 minutes before expiry
- Scope confirmed in production: `Calendars.ReadWrite`, `Mail.ReadWrite`, `Contacts.ReadWrite`, `Files.ReadWrite.All`, `Chat.Read`, and ~60 more
- The chrome-devtools-mcp connection and its underlying browser are kept alive across token acquisitions rather than launched fresh each time (`TokenManager.close()` tears down the connection and browser). The Chrome window is visible (never headless) from the moment the browser first starts — not just on failure. If no matching request shows up within 60s of navigating (polled every 5s), the server throws, telling the user to sign in in that already-open window and retry — because the connection stays alive, the retry reuses the same browser/page instead of opening a new window, so sign-in (including MFA) can be completed at the user's own pace across multiple retries.

## Key files

| File | Purpose |
|------|---------|
| `src/auth.ts` | `TokenManager` — drives Chrome via chrome-devtools-mcp, extracts the Bearer token from network requests, caches with expiry |
| `src/calendar.ts` | `CalendarClient` — calls `outlook.office.com/api/v2.0/me/calendarview` |
| `src/mail.ts` | `MailClient` — calls `outlook.office.com/api/v2.0` mail endpoints |
| `src/types.ts` | Shared types: `OwaToken`, `CalendarEvent`, `MailMessage`, payload interfaces |
| `src/index.ts` | MCP server, tool registration via `@modelcontextprotocol/sdk` |

## Development

```bash
npm install
npm run dev          # run with tsx (no build needed)
npm run build        # compile to dist/
npm test             # integration tests (require chrome-devtools-mcp + a live M365 session)
```

Tests are integration tests — they need a live M365 session in the chrome-devtools-mcp-driven Chrome profile. There are no unit tests with mocks because the auth flow is inherently side-effectful.

`npm test` runs Jest with `--runInBand`. Each test file gets its own `TokenManager`, and Jest runs files in separate worker processes by default — without `--runInBand`, multiple chrome-devtools-mcp instances launch concurrently against the same dedicated profile directory and lock each other out, causing every file but the first to spuriously fail with "no active session found."

## Adding new tools

1. Add any new API methods to `src/calendar.ts`, `src/mail.ts`, or create new domain clients (e.g., `src/contacts.ts`)
2. Register the tool in `src/index.ts` using `server.tool(name, description, zodSchema, handler)`
3. Add an integration test in `tests/`

**Tool output convention:** All tools must return structured JSON, not formatted text. Return the data array (or object) directly via `JSON.stringify(data, null, 2)`. The consuming LLM is better at reasoning over structured data and can format it however suits the user's question. Never pre-format, truncate fields, or drop metadata — pass through the full normalized object.

The OWA REST API base is `https://outlook.office.com/api/v2.0`. It mirrors the Microsoft Graph API shape closely — most Graph calendar/mail docs apply with `Subject`/`Start`/`End` casing instead of `subject`/`start`/`end`.

## Two API surfaces

This project uses **two different OWA API surfaces**:

### 1. REST API (`outlook.office.com/api/v2.0`)
Standard REST endpoints for CRUD operations. Used by calendar tools (`get_calendar_events`, `create_calendar_event`, `update_calendar_event`, `delete_calendar_event`) and all mail tools (`get_emails`, `search_emails`, `get_email`, `send_email`, `create_draft`, `create_reply_draft`, `create_reply_all_draft`, `create_forward_draft`, `update_draft`, `send_draft`, `move_email`, `delete_email`, `update_email`). Event IDs are in EwsId/RestId format (base64url, start with `AAMkA`).

### 2. OWA service.svc (`outlook.office.com/owa/service.svc`)
Internal OWA endpoint used by the Outlook Web client. Payload goes in the `x-owa-urlpostdata` header (URL-encoded JSON), not the request body (content-length=0). Used for `cancel_calendar_event` (all scopes), `follow_calendar_event`, and `respond_to_calendar_event`.

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

## Releasing

This package is published on npm as `owa-mcp`. A GitHub Actions workflow (`.github/workflows/publish.yml`) automates publishing on tag push using **npm Trusted Publishers** (OIDC — no token required):

1. **Bump version** in `package.json` (semver: patch for fixes, minor for new tools, major for breaking changes)
2. **Commit** the version bump: `git commit -am "chore: bump version to 0.X.0"`
3. **Tag** with `v` prefix: `git tag v0.X.0` (tag must match package.json version with `v` prefix)
4. **Push commit and tag**: `git push origin main --tags`
5. **Automated** (GitHub Actions): builds, publishes to npm via OIDC (no NPM_TOKEN needed), generates provenance attestation, creates a GitHub release with auto-generated notes

**npm auth**: Uses Trusted Publishers — GitHub Actions authenticates to npm via OIDC. No secrets to manage or rotate. Configured on npmjs.com → package settings → Publishing access → Trusted publishers. The `id-token: write` permission is required in the workflow.

When using `/autoresearch:ship`, select the `code-release` type (not "direct commit") to trigger the full release flow. A direct commit to main is for intermediate work; a release is for publishable milestones.

## Known limitations

- macOS only (the Chrome profile path, `~/Library/Application Support/owa-mcp/chrome-profile`, is hardcoded to a Mac location)
- Requires chrome-devtools-mcp to be runnable via `npx chrome-devtools-mcp@latest` and Chrome to be installed — there is no fallback if either is unavailable
- Token acquisition takes ~8–10 seconds on cold start (chrome-devtools-mcp spawn + browser launch); subsequent acquisitions reuse the already-running browser and are faster
- If the M365 session in the dedicated Chrome profile expires (usually after weeks of inactivity), the user must sign in to outlook.office.com in the already-open automation Chrome window, then retry the request
