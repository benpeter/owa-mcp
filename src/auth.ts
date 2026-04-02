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
