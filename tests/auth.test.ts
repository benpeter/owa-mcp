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
