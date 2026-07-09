// src/auth.ts
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';
import path from 'path';
import os from 'os';
import type { OwaToken } from './types.js';

// Dedicated, persistent Chrome profile for owa-mcp's automation browser.
// Kept separate from the user's daily-driver Chrome profile (and from any
// chrome-devtools-mcp instance a Claude Code session may already have open)
// to avoid user-data-dir lock conflicts.
const PROFILE_DIR = path.join(
  os.homedir(),
  'Library/Application Support/owa-mcp/chrome-profile'
);

const OWA_URL = 'https://outlook.office.com/calendar/view/workweek';

// Outlook Web makes OWA service calls with this token — it carries
// Calendars.ReadWrite and full Mail scope. The app ID in the token is
// 9199bf20-a13f-4107-85dc-02114787ef48 (Microsoft's OWA web app).
// outlook.office.com/calendar/view/... 302-redirects to outlook.cloud.microsoft
// for some tenants (Microsoft's ongoing domain consolidation), so the hostname
// alone can't be hardcoded to one value — but it must still be checked exactly
// against this allowlist (not a substring match) so a page reached during the
// interactive sign-in flow can't spoof a "/owa/service.svc" path on an
// attacker-controlled domain and get its request headers treated as ours.
const OWA_TOKEN_HOSTS = new Set([
  'outlook.office.com',
  'outlook.office365.com',
  'outlook.cloud.microsoft',
]);
const OWA_TOKEN_PATH_PREFIX = '/owa/service.svc';

function isOwaServiceSvcUrl(rawUrl: string): boolean {
  let url: URL;
  try {
    url = new URL(rawUrl);
  } catch {
    return false;
  }
  return OWA_TOKEN_HOSTS.has(url.hostname) && url.pathname.startsWith(OWA_TOKEN_PATH_PREFIX);
}

// Refresh 5 minutes before actual expiry
const REFRESH_BUFFER_MS = 5 * 60 * 1000;

const CONNECT_TIMEOUT_MS = 45_000;
const POLL_INTERVAL_MS = 5_000;
const POLL_TIMEOUT_MS = 60_000;

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function withTimeout<T>(promise: Promise<T>, ms: number, message: string): Promise<T> {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error(message)), ms);
    promise.then(
      (value) => {
        clearTimeout(timer);
        resolve(value);
      },
      (err) => {
        clearTimeout(timer);
        reject(err);
      }
    );
  });
}

export class TokenManager {
  private cached: OwaToken | null = null;
  private inflightPromise: Promise<OwaToken> | null = null;
  private client: Client | null = null;
  // True when the automation page needs a fresh navigation before polling
  // (first connection, or right after a successful acquisition — to force a
  // new request on the next refresh cycle). Left false across a "please log
  // in" timeout so retries don't interrupt an in-progress sign-in by
  // reloading the page out from under the user.
  private needsNavigation = true;

  /** Returns a valid token, refreshing automatically when near expiry. */
  async getToken(): Promise<OwaToken> {
    if (this.cached && this.isValid(this.cached)) {
      return this.cached;
    }
    // Coalesce concurrent callers into one acquisition
    if (!this.inflightPromise) {
      this.inflightPromise = this.acquireToken().finally(() => {
        this.inflightPromise = null;
      });
    }
    return this.inflightPromise;
  }

  /** Closes the chrome-devtools-mcp connection and its browser, if one is open. */
  async close(): Promise<void> {
    if (this.client) {
      const client = this.client;
      this.client = null;
      await client.close().catch(() => {});
    }
  }

  private isValid(token: OwaToken): boolean {
    return token.expiresAt - REFRESH_BUFFER_MS > Date.now();
  }

  /**
   * chrome-devtools-mcp is a hard prerequisite — this project has no other
   * way to reach an authenticated Outlook Web session. Spawns it on first
   * use via npx and keeps the connection alive across token acquisitions
   * (rather than relaunching per call) so a browser window opened for the
   * user to sign in stays open across retries.
   */
  private async ensureBrowserConnection(): Promise<Client> {
    if (this.client) return this.client;

    const transport = new StdioClientTransport({
      command: 'npx',
      args: [
        'chrome-devtools-mcp@latest',
        `--userDataDir=${PROFILE_DIR}`,
        '--viewport=1280x800',
      ],
      stderr: 'pipe',
    });
    let stderrOutput = '';
    transport.stderr?.on('data', (chunk: Buffer) => {
      stderrOutput = (stderrOutput + chunk.toString()).slice(-500);
    });

    const client = new Client({ name: 'owa-mcp', version: '1.0.0' });
    try {
      await withTimeout(
        client.connect(transport),
        CONNECT_TIMEOUT_MS,
        `chrome-devtools-mcp did not start within ${CONNECT_TIMEOUT_MS / 1000}s`
      );
    } catch (err) {
      await transport.close().catch(() => {});
      const detail = stderrOutput.trim();
      throw new Error(
        'chrome-devtools-mcp is required but could not be started' +
          (detail ? `: ${detail}` : ` (${(err as Error).message})`) +
          '. Verify it works with "npx chrome-devtools-mcp@latest --version" and that ' +
          'Chrome is installed, then retry. See https://github.com/ChromeDevTools/chrome-devtools-mcp for setup.'
      );
    }

    client.onclose = () => {
      if (this.client === client) {
        this.client = null;
        this.needsNavigation = true;
      }
    };
    this.client = client;
    this.needsNavigation = true;
    return client;
  }

  private async callToolText(
    client: Client,
    name: string,
    args: Record<string, unknown>
  ): Promise<string> {
    const result = await client.callTool({ name, arguments: args });
    const content = (result.content ?? []) as Array<{ type: string; text?: string }>;
    return content
      .filter((c) => c.type === 'text')
      .map((c) => c.text ?? '')
      .join('\n');
  }

  // Returns all matching reqids (oldest first) plus the total request count
  // seen, for diagnostics if none ever match.
  // No resourceTypes filter: OWA's service.svc XHR/fetch calls aren't always
  // classified as 'xhr'/'fetch' by Puppeteer's CDP-derived resourceType, and
  // isOwaServiceSvcUrl() already narrows precisely — a resourceType filter
  // here only risks silently excluding a real match, never helps.
  private async findServiceSvcRequestIds(
    client: Client
  ): Promise<{ ids: number[]; totalSeen: number }> {
    const listing = await this.callToolText(client, 'list_network_requests', {
      includePreservedRequests: true,
    });
    const ids: number[] = [];
    let totalSeen = 0;
    for (const line of listing.split('\n')) {
      const match = line.match(/^reqid=(\d+)\s+\S+\s+(\S+)\s+\[/);
      if (!match) continue;
      totalSeen++;
      if (isOwaServiceSvcUrl(match[2])) {
        ids.push(Number(match[1]));
      }
    }
    return { ids, totalSeen };
  }

  private async extractBearerToken(client: Client, reqid: number): Promise<string | null> {
    const detail = await this.callToolText(client, 'get_network_request', { reqid });
    const match = detail.match(/^-\s*authorization\s*:\s*(.+)$/im);
    if (!match) return null;
    return match[1].trim().replace(/^Bearer\s+/i, '').trim();
  }

  private async acquireToken(): Promise<OwaToken> {
    const client = await this.ensureBrowserConnection();

    if (this.needsNavigation) {
      await this.callToolText(client, 'navigate_page', {
        type: 'url',
        url: OWA_URL,
        timeout: 30_000,
      });
      this.needsNavigation = false;
    }

    const deadline = Date.now() + POLL_TIMEOUT_MS;
    let lastTotalSeen = 0;
    let lastMatchCount = 0;
    while (true) {
      // Try newest-first: an earlier service.svc call (e.g. an early
      // bootstrap/ping before OWA has attached its Bearer token) can
      // permanently lack a usable Authorization header, so we can't just
      // take the first match — fall through to older ones if a newer one
      // has no valid header either.
      const { ids: reqids, totalSeen } = await this.findServiceSvcRequestIds(client);
      lastTotalSeen = totalSeen;
      lastMatchCount = reqids.length;
      for (let i = reqids.length - 1; i >= 0; i--) {
        const raw = await this.extractBearerToken(client, reqids[i]);
        if (raw) {
          const token = this.parseToken(raw);
          this.cached = token;
          // Force a fresh navigation on the next refresh cycle so a new
          // service.svc request actually fires instead of relying on
          // OWA's own background silent-renewal timing.
          this.needsNavigation = true;
          return token;
        }
      }
      if (Date.now() >= deadline) {
        throw new Error(
          'No active Microsoft 365 session was found in the automation browser after ' +
            `${POLL_TIMEOUT_MS / 1000}s (saw ${lastTotalSeen} network requests, ` +
            `${lastMatchCount} matching ${OWA_TOKEN_PATH_PREFIX} but none had a usable ` +
            'Authorization header). A Chrome window has been opened — please complete ' +
            'sign-in there, then retry this request.'
        );
      }
      await sleep(POLL_INTERVAL_MS);
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
