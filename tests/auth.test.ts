// tests/auth.test.ts
import { TokenManager } from '../src/auth.js';

// Integration test — requires chrome-devtools-mcp to be available via npx,
// Chrome to be installed, and an active (or completable) Microsoft 365
// session in the dedicated automation Chrome profile at
// ~/Library/Application Support/owa-mcp/chrome-profile.
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
  }, 150_000); // chrome-devtools-mcp startup (45s) + navigate_page (30s) + session poll (60s)

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
