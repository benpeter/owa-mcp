# OWA MCP Foundation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a working MCP server that authenticates against Microsoft 365 Outlook Web App via a locally signed-in Microsoft Edge browser profile and exposes a `get_calendar_events` tool to Claude Code.

**Architecture:** The server intercepts Bearer tokens emitted by Outlook Web (outlook.office.com) when it loads in a headless Playwright-controlled Edge browser that uses the user's existing Edge profile. Those tokens are valid for `outlook.office.com/api/v2.0` REST endpoints with full `Calendars.ReadWrite` scope — no Azure app registration required. Tokens are cached in memory and refreshed automatically (they expire in ~80 minutes) by re-launching the headless browser.

**Tech Stack:** Node.js 20+, TypeScript, `@modelcontextprotocol/sdk`, `playwright` (msedge channel), `zod`, `tsx` for dev, `tsup` for build.

---

## Background: How Auth Works

This is critical context — do not skip.

New Outlook on Mac dropped AppleScript. Azure CLI auth to corporate M365 tenants is blocked by Conditional Access policies on managed devices. Registering an Azure app requires IT admin access in enterprise tenants.

The solution: Microsoft Edge on Mac, when signed in to Microsoft 365, holds a valid MSAL session. When Playwright launches Edge with `--user-data-dir` pointing at the real Edge profile directory (`~/Library/Application Support/Microsoft Edge`), the browser is already authenticated. Outlook Web makes API calls using Bearer tokens issued to the OWA app ID (`9199bf20-a13f-4107-85dc-02114787ef48`). These tokens carry extensive delegated scopes including `Calendars.ReadWrite`, `Mail.ReadWrite`, etc.

The MCP server intercepts these tokens via Playwright's `page.on('request', ...)` handler, then uses them directly against `https://outlook.office.com/api/v2.0` — a proper JSON REST API (same as Graph but hosted on the Outlook endpoint).

**Token lifetime:** ~80 minutes from issue. The server tracks `exp` from the JWT payload and auto-refreshes before expiry.

**Headless works:** Confirmed. The Edge profile session is valid in headless mode; the token can be acquired without a visible browser window (~8–10s cold start, then cached).

---

## File Structure

```
owa-mcp/
├── src/
│   ├── index.ts          # MCP server entrypoint, tool registration
│   ├── auth.ts           # Token acquisition via Playwright + Edge profile
│   ├── calendar.ts       # Calendar API calls using OWA REST API
│   └── types.ts          # Shared TypeScript types
├── tests/
│   ├── auth.test.ts      # Token acquisition tests (integration, needs Edge)
│   └── calendar.test.ts  # Calendar API tests (integration, needs valid token)
├── CLAUDE.md             # Context for future Claude Code sessions
├── README.md             # Public documentation
├── LICENSE               # Apache 2.0
├── package.json
├── tsconfig.json
└── .gitignore
```

---

## Task 1: Project Scaffold

**Files:**
- Create: `package.json`
- Create: `tsconfig.json`
- Create: `.gitignore`
- Create: `src/types.ts`

- [ ] **Step 1: Initialize git repo and create package.json**

```bash
cd /Users/I752296/github/benpeter/owa-mcp
git init
cat > package.json << 'EOF'
{
  "name": "owa-mcp",
  "version": "0.1.0",
  "description": "MCP server for Microsoft Outlook calendar via Playwright Edge session interception",
  "type": "module",
  "main": "dist/index.js",
  "bin": {
    "owa-mcp": "dist/index.js"
  },
  "scripts": {
    "dev": "tsx src/index.ts",
    "build": "tsup src/index.ts --format esm --dts --out-dir dist",
    "test": "node --experimental-vm-modules node_modules/.bin/jest"
  },
  "dependencies": {
    "@modelcontextprotocol/sdk": "^1.10.0",
    "playwright": "^1.51.0",
    "zod": "^3.24.0"
  },
  "devDependencies": {
    "@types/node": "^22.0.0",
    "jest": "^29.7.0",
    "ts-jest": "^29.3.0",
    "tsup": "^8.4.0",
    "tsx": "^4.19.0",
    "typescript": "^5.8.0"
  },
  "engines": {
    "node": ">=20.0.0"
  }
}
EOF
```

- [ ] **Step 2: Create tsconfig.json**

```bash
cat > tsconfig.json << 'EOF'
{
  "compilerOptions": {
    "target": "ES2022",
    "module": "ESNext",
    "moduleResolution": "bundler",
    "strict": true,
    "outDir": "dist",
    "rootDir": "src",
    "declaration": true,
    "skipLibCheck": true,
    "esModuleInterop": true
  },
  "include": ["src/**/*"],
  "exclude": ["node_modules", "dist", "tests"]
}
EOF
```

- [ ] **Step 3: Create .gitignore**

```bash
cat > .gitignore << 'EOF'
node_modules/
dist/
*.js.map
.env
EOF
```

- [ ] **Step 4: Create src/types.ts**

```typescript
// src/types.ts
// tva

export interface OwaToken {
  value: string;       // raw JWT
  expiresAt: number;   // unix epoch ms
  issuedAt: number;    // unix epoch ms
}

export interface CalendarEvent {
  id: string;
  subject: string;
  start: string;       // ISO 8601
  end: string;         // ISO 8601
  isAllDay: boolean;
  organizer: string;
  location: string;
  isOnlineMeeting: boolean;
  showAs: string;      // Free | Tentative | Busy | Oof | WorkingElsewhere | Unknown
  isRecurring: boolean;
  isPrivate: boolean;
  bodyPreview: string;
}

export interface OwaCalendarViewResponse {
  value: OwaCalendarEvent[];
  '@odata.nextLink'?: string;
}

// Raw shape returned by outlook.office.com/api/v2.0/me/calendarview
export interface OwaCalendarEvent {
  Id: string;
  Subject: string;
  Start: { DateTime: string; TimeZone: string };
  End: { DateTime: string; TimeZone: string };
  IsAllDay: boolean;
  Organizer: { EmailAddress: { Name: string; Address: string } };
  Location: { DisplayName: string };
  IsOnlineMeeting: boolean;
  ShowAs: string;
  IsReminderOn: boolean;
  Recurrence: unknown | null;
  Sensitivity: string;   // Normal | Personal | Private | Confidential
  BodyPreview: string;
}
```

- [ ] **Step 5: Install dependencies**

```bash
npm install
```

Expected output: `added N packages`

- [ ] **Step 6: Commit scaffold**

```bash
git add .
git commit -m "feat: project scaffold — types, tsconfig, package.json"
```

---

## Task 2: Token Acquisition (`src/auth.ts`)

**Files:**
- Create: `src/auth.ts`
- Create: `tests/auth.test.ts`

- [ ] **Step 1: Write the failing test**

```bash
mkdir -p tests
cat > tests/auth.test.ts << 'EOF'
// tests/auth.test.ts
import { TokenManager } from '../src/auth.js';

// Integration test — requires Microsoft Edge installed at default path
// and an active M365 session in the Edge profile.
// Run manually: npm test -- --testPathPattern=auth

describe('TokenManager', () => {
  let manager: TokenManager;

  beforeAll(() => {
    manager = new TokenManager();
  });

  afterAll(async () => {
    await manager.close();
  });

  test('acquires a Bearer token from Outlook Web', async () => {
    const token = await manager.getToken();
    expect(token.value).toMatch(/^eyJ/);           // JWT starts with eyJ
    expect(token.expiresAt).toBeGreaterThan(Date.now());
    expect(token.expiresAt - token.issuedAt).toBeGreaterThan(60 * 60 * 1000); // >1hr
  }, 30_000);

  test('returns cached token on second call', async () => {
    const t1 = await manager.getToken();
    const t2 = await manager.getToken();
    expect(t1.value).toBe(t2.value);
  }, 5_000);

  test('token is valid for OWA REST API', async () => {
    const token = await manager.getToken();
    const res = await fetch('https://outlook.office.com/api/v2.0/me', {
      headers: { Authorization: `Bearer ${token.value}` }
    });
    expect(res.status).toBe(200);
    const data = await res.json() as { EmailAddress: string };
    expect(data.EmailAddress).toMatch(/@/);
  }, 10_000);
});
EOF
```

- [ ] **Step 2: Add jest config to package.json**

Edit `package.json` — replace the `"scripts"` and add `"jest"` config:

```json
{
  "scripts": {
    "dev": "tsx src/index.ts",
    "build": "tsup src/index.ts --format esm --dts --out-dir dist",
    "test": "node --experimental-vm-modules node_modules/.bin/jest --testTimeout=30000"
  },
  "jest": {
    "preset": "ts-jest/presets/default-esm",
    "testEnvironment": "node",
    "extensionsToTreatAsEsm": [".ts"],
    "moduleNameMapper": {
      "^(\\.{1,2}/.*)\\.js$": "$1"
    },
    "transform": {
      "^.+\\.tsx?$": ["ts-jest", { "useESM": true }]
    },
    "testPathPattern": "tests/"
  }
}
```

- [ ] **Step 3: Run the test — verify it fails**

```bash
npm test -- --testPathPattern=auth 2>&1 | tail -10
```

Expected: FAIL with `Cannot find module '../src/auth.js'`

- [ ] **Step 4: Implement src/auth.ts**

```typescript
// src/auth.ts
import { chromium, type BrowserContext } from 'playwright';
import path from 'path';
import os from 'os';
import type { OwaToken } from './types.js';

const EDGE_PROFILE_DIR = path.join(
  os.homedir(),
  'Library/Application Support/Microsoft Edge'
);

// Outlook Web makes OWA service calls with this token — it carries
// Calendars.ReadWrite and full Mail scope. The app ID in the token is
// 9199bf20-a13f-4107-85dc-02114787ef48 (Microsoft's OWA web app).
const OWA_TOKEN_URL_PATTERN = 'outlook.office.com/owa/service.svc';

// Refresh 5 minutes before actual expiry
const REFRESH_BUFFER_MS = 5 * 60 * 1000;

export class TokenManager {
  private cached: OwaToken | null = null;
  private inflightPromise: Promise<OwaToken> | null = null;

  /** Returns a valid token, refreshing automatically when near expiry. */
  async getToken(): Promise<OwaToken> {
    if (this.cached && this.isValid(this.cached)) {
      return this.cached;
    }
    // Coalesce concurrent callers into one browser launch
    if (!this.inflightPromise) {
      this.inflightPromise = this.acquireToken().finally(() => {
        this.inflightPromise = null;
      });
    }
    return this.inflightPromise;
  }

  /** No-op: TokenManager is stateless between acquisitions (no persistent browser). */
  async close(): Promise<void> {
    // Nothing to clean up — each acquisition opens and closes its own browser.
  }

  private isValid(token: OwaToken): boolean {
    return token.expiresAt - REFRESH_BUFFER_MS > Date.now();
  }

  private async acquireToken(): Promise<OwaToken> {
    let context: BrowserContext | null = null;
    try {
      context = await chromium.launchPersistentContext(EDGE_PROFILE_DIR, {
        channel: 'msedge',
        headless: true,
        args: ['--no-first-run', '--no-default-browser-check'],
      });

      const page = await context.newPage();
      const tokenPromise = new Promise<string>((resolve, reject) => {
        const timeout = setTimeout(
          () => reject(new Error('Timed out waiting for OWA Bearer token (25s)')),
          25_000
        );
        page.on('request', (req) => {
          const auth = req.headers()['authorization'];
          if (auth && req.url().includes(OWA_TOKEN_URL_PATTERN)) {
            clearTimeout(timeout);
            resolve(auth.replace(/^Bearer\s+/i, '').trim());
          }
        });
      });

      await page.goto('https://outlook.office.com/calendar/view/workweek', {
        waitUntil: 'domcontentloaded',
        timeout: 30_000,
      });

      const rawToken = await tokenPromise;
      const token = this.parseToken(rawToken);
      this.cached = token;
      return token;
    } finally {
      await context?.close();
    }
  }

  private parseToken(raw: string): OwaToken {
    const parts = raw.split('.');
    if (parts.length !== 3) throw new Error('Invalid JWT structure');
    const payload = JSON.parse(
      Buffer.from(parts[1], 'base64url').toString('utf8')
    ) as { exp: number; iat: number };
    return {
      value: raw,
      expiresAt: payload.exp * 1000,
      issuedAt: payload.iat * 1000,
    };
  }
}
```

- [ ] **Step 5: Run the tests — verify they pass**

```bash
npm test -- --testPathPattern=auth 2>&1 | tail -20
```

Expected: all 3 tests PASS. If token acquisition times out, verify Edge is installed and signed in to M365.

- [ ] **Step 6: Commit**

```bash
git add src/auth.ts tests/auth.test.ts package.json
git commit -m "feat: TokenManager — acquire OWA Bearer token via headless Edge"
```

---

## Task 3: Calendar API (`src/calendar.ts`)

**Files:**
- Create: `src/calendar.ts`
- Create: `tests/calendar.test.ts`

- [ ] **Step 1: Write the failing test**

```bash
cat > tests/calendar.test.ts << 'EOF'
// tests/calendar.test.ts
import { TokenManager } from '../src/auth.js';
import { CalendarClient } from '../src/calendar.js';

describe('CalendarClient', () => {
  let manager: TokenManager;
  let client: CalendarClient;

  beforeAll(async () => {
    manager = new TokenManager();
    client = new CalendarClient(manager);
  });

  afterAll(async () => {
    await manager.close();
  });

  test('returns events for a date range', async () => {
    const start = new Date();
    const end = new Date(Date.now() + 7 * 24 * 60 * 60 * 1000);
    const events = await client.getCalendarEvents(start.toISOString(), end.toISOString());

    expect(Array.isArray(events)).toBe(true);
    expect(events.length).toBeGreaterThan(0);

    const first = events[0];
    expect(typeof first.id).toBe('string');
    expect(typeof first.subject).toBe('string');
    expect(first.start).toMatch(/^\d{4}-\d{2}-\d{2}T/);
    expect(first.end).toMatch(/^\d{4}-\d{2}-\d{2}T/);
    expect(typeof first.organizer).toBe('string');
  }, 40_000);

  test('handles empty range gracefully', async () => {
    // Far future range with no events
    const start = '2099-01-01T00:00:00Z';
    const end = '2099-01-02T00:00:00Z';
    const events = await client.getCalendarEvents(start, end);
    expect(Array.isArray(events)).toBe(true);
    expect(events.length).toBe(0);
  }, 40_000);

  test('respects maxResults parameter', async () => {
    const start = new Date();
    const end = new Date(Date.now() + 30 * 24 * 60 * 60 * 1000);
    const events = await client.getCalendarEvents(start.toISOString(), end.toISOString(), { maxResults: 3 });
    expect(events.length).toBeLessThanOrEqual(3);
  }, 40_000);
});
EOF
```

- [ ] **Step 2: Run the test — verify it fails**

```bash
npm test -- --testPathPattern=calendar 2>&1 | tail -10
```

Expected: FAIL with `Cannot find module '../src/calendar.js'`

- [ ] **Step 3: Implement src/calendar.ts**

```typescript
// src/calendar.ts
import type { TokenManager } from './auth.js';
import type { CalendarEvent, OwaCalendarViewResponse, OwaCalendarEvent } from './types.js';

const OWA_BASE = 'https://outlook.office.com/api/v2.0';

export interface GetCalendarEventsOptions {
  maxResults?: number;          // default 50
  timezone?: string;            // IANA tz name, default 'UTC'
}

export class CalendarClient {
  constructor(private readonly tokens: TokenManager) {}

  /**
   * Returns calendar events between startDateTime and endDateTime (ISO 8601 strings).
   * Handles OData paging automatically up to maxResults.
   */
  async getCalendarEvents(
    startDateTime: string,
    endDateTime: string,
    options: GetCalendarEventsOptions = {}
  ): Promise<CalendarEvent[]> {
    const { maxResults = 50, timezone = 'UTC' } = options;

    const token = await this.tokens.getToken();
    const params = new URLSearchParams({
      startDateTime,
      endDateTime,
      '$select': 'Id,Subject,Start,End,IsAllDay,Organizer,Location,IsOnlineMeeting,ShowAs,Recurrence,Sensitivity,BodyPreview',
      '$top': String(Math.min(maxResults, 100)),
      '$orderby': 'Start/DateTime asc',
    });

    const url = `${OWA_BASE}/me/calendarview?${params}`;
    const events: CalendarEvent[] = [];
    let nextLink: string | undefined = url;

    while (nextLink && events.length < maxResults) {
      const res = await fetch(nextLink, {
        headers: {
          Authorization: `Bearer ${token.value}`,
          Accept: 'application/json',
          Prefer: `outlook.timezone="${timezone}"`,
        },
      });

      if (!res.ok) {
        const body = await res.text();
        throw new Error(`OWA calendar API error ${res.status}: ${body}`);
      }

      const data = (await res.json()) as OwaCalendarViewResponse;
      for (const raw of data.value) {
        events.push(this.normalise(raw));
        if (events.length >= maxResults) break;
      }
      nextLink = data['@odata.nextLink'];
    }

    return events;
  }

  private normalise(raw: OwaCalendarEvent): CalendarEvent {
    return {
      id: raw.Id,
      subject: raw.Subject,
      start: raw.Start.DateTime,
      end: raw.End.DateTime,
      isAllDay: raw.IsAllDay,
      organizer: raw.Organizer?.EmailAddress?.Name ?? '',
      location: raw.Location?.DisplayName ?? '',
      isOnlineMeeting: raw.IsOnlineMeeting,
      showAs: raw.ShowAs,
      isRecurring: raw.Recurrence !== null,
      isPrivate: raw.Sensitivity === 'Private',
      bodyPreview: raw.BodyPreview ?? '',
    };
  }
}
```

- [ ] **Step 4: Run the tests — verify they pass**

```bash
npm test -- --testPathPattern=calendar 2>&1 | tail -20
```

Expected: all 3 tests PASS.

- [ ] **Step 5: Commit**

```bash
git add src/calendar.ts tests/calendar.test.ts
git commit -m "feat: CalendarClient — getCalendarEvents via OWA REST API v2.0"
```

---

## Task 4: MCP Server Entrypoint (`src/index.ts`)

**Files:**
- Create: `src/index.ts`

- [ ] **Step 1: Install Playwright browser (msedge channel — must be system Edge)**

Note: `playwright install msedge` downloads a Playwright-managed Chromium. For this project we need the system Edge at the default Mac path. The `channel: 'msedge'` option in Playwright uses the system-installed Edge, not a downloaded binary — no install needed. Verify Edge is present:

```bash
ls "/Applications/Microsoft Edge.app" && echo "Edge found"
```

Expected: `Edge found`

- [ ] **Step 2: Implement src/index.ts**

```typescript
#!/usr/bin/env node
// src/index.ts
// tva

import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { z } from 'zod';
import { TokenManager } from './auth.js';
import { CalendarClient } from './calendar.js';

const tokenManager = new TokenManager();
const calendarClient = new CalendarClient(tokenManager);

const server = new McpServer({
  name: 'owa-mcp',
  version: '0.1.0',
});

server.tool(
  'get_calendar_events',
  'Retrieve calendar events from Microsoft Outlook. Returns events between startDateTime and endDateTime.',
  {
    startDateTime: z
      .string()
      .describe('Start of time range in ISO 8601 format, e.g. 2026-04-07T00:00:00Z'),
    endDateTime: z
      .string()
      .describe('End of time range in ISO 8601 format, e.g. 2026-04-14T00:00:00Z'),
    maxResults: z
      .number()
      .int()
      .min(1)
      .max(100)
      .optional()
      .default(50)
      .describe('Maximum number of events to return (default 50, max 100)'),
    timezone: z
      .string()
      .optional()
      .default('UTC')
      .describe('IANA timezone name for event times, e.g. Europe/Berlin'),
  },
  async ({ startDateTime, endDateTime, maxResults, timezone }) => {
    const events = await calendarClient.getCalendarEvents(startDateTime, endDateTime, {
      maxResults,
      timezone,
    });

    if (events.length === 0) {
      return {
        content: [{ type: 'text', text: 'No events found in the specified time range.' }],
      };
    }

    const lines = events.map((e) => {
      const time = e.isAllDay
        ? `All day`
        : `${e.start} → ${e.end}`;
      const flags = [
        e.isOnlineMeeting ? 'Teams' : '',
        e.isRecurring ? 'Recurring' : '',
        e.isPrivate ? 'Private' : '',
        e.showAs !== 'Busy' ? e.showAs : '',
      ]
        .filter(Boolean)
        .join(', ');
      return [
        `**${e.subject}**`,
        `  Time: ${time}`,
        e.organizer ? `  Organizer: ${e.organizer}` : '',
        e.location ? `  Location: ${e.location}` : '',
        flags ? `  Flags: ${flags}` : '',
        e.bodyPreview ? `  Preview: ${e.bodyPreview.substring(0, 120)}` : '',
      ]
        .filter(Boolean)
        .join('\n');
    });

    return {
      content: [
        {
          type: 'text',
          text: `Found ${events.length} event(s):\n\n${lines.join('\n\n')}`,
        },
      ],
    };
  }
);

async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  // stderr so it doesn't interfere with MCP stdio protocol
  process.stderr.write('owa-mcp server running on stdio\n');
}

main().catch((err) => {
  process.stderr.write(`Fatal: ${err}\n`);
  process.exit(1);
});
```

- [ ] **Step 3: Build and do a quick smoke test**

```bash
npm run build 2>&1 | tail -10
```

Expected: `dist/index.js` created with no errors.

```bash
# Smoke test: list tools via MCP initialize + tools/list
echo '{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2024-11-05","capabilities":{},"clientInfo":{"name":"test","version":"1"}}}
{"jsonrpc":"2.0","id":2,"method":"tools/list","params":{}}' | node dist/index.js 2>/dev/null | head -5
```

Expected: JSON response containing `"get_calendar_events"`.

- [ ] **Step 4: Commit**

```bash
git add src/index.ts
git commit -m "feat: MCP server entrypoint with get_calendar_events tool"
```

---

## Task 5: Documentation (README, CLAUDE.md, LICENSE)

**Files:**
- Create: `README.md`
- Create: `CLAUDE.md`
- Create: `LICENSE`

- [ ] **Step 1: Write README.md**

```markdown
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
```

- [ ] **Step 2: Write CLAUDE.md**

```markdown
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
```

- [ ] **Step 3: Write LICENSE**

```
                                 Apache License
                           Version 2.0, January 2004
                        http://www.apache.org/licenses/

   TERMS AND CONDITIONS FOR USE, REPRODUCTION, AND DISTRIBUTION

   1. Definitions.

      "License" shall mean the terms and conditions for use, reproduction,
      and distribution as defined by Sections 1 through 9 of this document.

      "Licensor" shall mean the copyright owner or entity authorized by
      the copyright owner that is granting the License.

      "Legal Entity" shall mean the union of the acting entity and all
      other entities that control, are controlled by, or are under common
      control with that entity. For the purposes of this definition,
      "control" means (i) the power, direct or indirect, to cause the
      direction or management of such entity, whether by contract or
      otherwise, or (ii) ownership of fifty percent (50%) or more of the
      outstanding shares, or (iii) beneficial ownership of such entity.

      "You" (or "Your") shall mean an individual or Legal Entity
      exercising permissions granted by this License.

      "Source" form shall mean the preferred form for making modifications,
      including but not limited to software source code, documentation
      source, and configuration files.

      "Object" form shall mean any form resulting from mechanical
      transformation or translation of a Source form, including but
      not limited to compiled object code, generated documentation,
      and conversions to other media types.

      "Work" shall mean the work of authorship made available under
      the License, as indicated by a copyright notice that is included in
      or attached to the work (an example is provided in the Appendix below).

      "Derivative Works" shall mean any work, whether in Source or Object
      form, that is based on (or derived from) the Work and for which the
      editorial revisions, annotations, elaborations, or other transformations
      represent, as a whole, an original work of authorship. For the purposes
      of this License, Derivative Works shall not include works that remain
      separable from, or merely link (or bind by name) to the interfaces of,
      the Work and Derivative Works thereof.

      "Contribution" shall mean, as submitted to the Licensor for inclusion
      in the Work by the copyright owner or by an individual or Legal Entity
      authorized to submit on behalf of the copyright owner. For the purposes
      of this definition, "submitted" means any form of electronic, verbal,
      or written communication sent to the Licensor or its representatives,
      including but not limited to communication on electronic mailing lists,
      source code control systems, and issue tracking systems that are managed
      by, or on behalf of, the Licensor for the purpose of recording and
      discussing the Work, but excluding communication that is conspicuously
      marked or designated in writing by the copyright owner as "Not a
      Contribution."

      "Contributor" shall mean Licensor and any Legal Entity on behalf of
      whom a Contribution has been received by the Licensor and included
      within the Work.

   2. Grant of Copyright License. Subject to the terms and conditions of
      this License, each Contributor hereby grants to You a perpetual,
      worldwide, non-exclusive, no-charge, royalty-free, irrevocable
      copyright license to reproduce, prepare Derivative Works of,
      publicly display, publicly perform, sublicense, and distribute the
      Work and such Derivative Works in Source or Object form.

   3. Grant of Patent License. Subject to the terms and conditions of
      this License, each Contributor hereby grants to You a perpetual,
      worldwide, non-exclusive, no-charge, royalty-free, irrevocable
      (except as stated in this section) patent license to make, have made,
      use, offer to sell, sell, import, and otherwise transfer the Work,
      where such license applies only to those patent claims licensable
      by such Contributor that are necessarily infringed by their
      Contribution(s) alone or by the combined work (in which their
      Contribution(s) with the Work. If You institute patent litigation
      against any entity (including a cross-claim or counterclaim in a
      lawsuit) alleging that the Work or any Contribution embodied within
      the Work constitutes direct or contributory patent infringement,
      then any patent rights granted to You under this License for that
      Work shall terminate as of the date such litigation is filed.

   4. Redistribution. You may reproduce and distribute copies of the
      Work or Derivative Works thereof in any medium, with or without
      modifications, and in Source or Object form, provided that You
      meet the following conditions:

      (a) You must give any other recipients of the Work or Derivative
          Works a copy of this License; and

      (b) You must cause any modified files to carry prominent notices
          stating that You changed the files; and

      (c) You must retain, in the Source form of any Derivative Works
          that You distribute, all copyright, patent, trademark, and
          attribution notices from the Source form of the Work,
          excluding those notices that do not pertain to any part of
          the Derivative Works; and

      (d) If the Work includes a "NOTICE" text file as part of its
          distribution, You must include a readable copy of the
          attribution notices contained within such NOTICE file, in
          at least one of the following places: within a NOTICE text
          file distributed as part of the Derivative Works; within
          the Source form or documentation, if provided along with the
          Derivative Works; or, within a display generated by the
          Derivative Works, if and wherever such third-party notices
          normally appear. The contents of the NOTICE file are for
          informational purposes only and do not modify the License.
          You may add Your own attribution notices within Derivative
          Works that You distribute, alongside or in addition to the
          NOTICE text from the Work, provided that such additional
          attribution notices cannot be construed as modifying the License.

   You may add Your own license statement for Your modifications and
   may provide additional terms or conditions for use, reproduction,
   or distribution of Your modifications, or for such Derivative Works
   as a whole, provided Your use, reproduction, and distribution of
   the Work otherwise complies with the conditions stated in this License.

   5. Submission of Contributions. Unless You explicitly state otherwise,
      any Contribution intentionally submitted for inclusion in the Work
      by You to the Licensor shall be under the terms and conditions of
      this License, without any additional terms or conditions.
      Notwithstanding the above, nothing herein shall supersede or modify
      the terms of any separate license agreement you may have executed
      with Licensor regarding such Contributions.

   6. Trademarks. This License does not grant permission to use the trade
      names, trademarks, service marks, or product names of the Licensor,
      except as required for reasonable and customary use in describing the
      origin of the Work and reproducing the content of the NOTICE file.

   7. Disclaimer of Warranty. Unless required by applicable law or
      agreed to in writing, Licensor provides the Work (and each
      Contributor provides its Contributions) on an "AS IS" BASIS,
      WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or
      implied, including, without limitation, any conditions of TITLE,
      NONINFRINGEMENT, MERCHANTABILITY, or FITNESS FOR A PARTICULAR PURPOSE.
      You are solely responsible for determining the appropriateness of
      using or reproducing the Work and assume any risks associated with
      Your exercise of permissions under this License.

   8. Limitation of Liability. In no event and under no legal theory,
      whether in tort (including negligence), contract, or otherwise,
      unless required by applicable law (such as deliberate and grossly
      negligent acts) or agreed to in writing, shall any Contributor be
      liable to You for damages, including any direct, indirect, special,
      incidental, or exemplary damages of any character arising as a
      result of this License or out of the use or inability to use the
      Work (including but not limited to damages for loss of goodwill,
      work stoppage, computer failure or malfunction, or all other
      commercial damages or losses), even if such Contributor has been
      advised of the possibility of such damages.

   9. Accepting Warranty or Additional Liability. While redistributing
      the Work or Derivative Works thereof, You may choose to offer,
      and charge a fee for, acceptance of support, warranty, indemnity,
      or other liability obligations and/or rights consistent with this
      License. However, in accepting such obligations, You may offer only
      conditions consistent with this License.

   END OF TERMS AND CONDITIONS

   Copyright 2026 Ben Peter

   Licensed under the Apache License, Version 2.0 (the "License");
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
```

- [ ] **Step 4: Commit docs**

```bash
git add README.md CLAUDE.md LICENSE
git commit -m "docs: README, CLAUDE.md context, Apache 2.0 license"
```

---

## Task 6: Create GitHub Repo and Push

**Files:** none (git operations only)

- [ ] **Step 1: Create public GitHub repo**

```bash
gh repo create benpeter/owa-mcp \
  --public \
  --description "MCP server for Microsoft Outlook calendar via Edge session token interception — no Azure app registration required" \
  --source . \
  --remote origin \
  --push
```

Expected: repo created at `https://github.com/benpeter/owa-mcp` and all commits pushed.

- [ ] **Step 2: Verify**

```bash
gh repo view benpeter/owa-mcp --web 2>/dev/null || echo "Open https://github.com/benpeter/owa-mcp"
```

---

## Task 7: Wire into Claude Code

**Files:** `~/.claude/settings.json` (modify)

- [ ] **Step 1: Add MCP server to Claude Code settings**

Read current settings, add the `owa` server:

```json
{
  "mcpServers": {
    "owa": {
      "command": "node",
      "args": ["/Users/I752296/github/benpeter/owa-mcp/dist/index.js"]
    }
  }
}
```

- [ ] **Step 2: Verify tool is available**

Restart Claude Code (or run `/mcp` to reload). Confirm `get_calendar_events` appears in the tool list.

---

## Self-Review

**Spec coverage:**
- Token acquisition via headless Edge: Task 2 ✓
- `get_calendar_events` MCP tool: Tasks 3 + 4 ✓
- README (public, no corporate mentions): Task 5 ✓
- CLAUDE.md with full context: Task 5 ✓
- Apache 2.0 license: Task 5 ✓
- GitHub public repo: Task 6 ✓
- Wired into Claude Code settings: Task 7 ✓

**Placeholder scan:** None found — all code steps are complete.

**Type consistency:** `OwaToken`, `CalendarEvent`, `OwaCalendarEvent` defined in `types.ts` Task 1 and referenced consistently in Tasks 2, 3, 4.
