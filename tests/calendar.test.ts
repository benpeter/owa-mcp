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
    // Very short range in far future — no events expected
    const start = '2099-12-31T23:00:00Z';
    const end = '2099-12-31T23:30:00Z';
    const events = await client.getCalendarEvents(start, end);
    expect(Array.isArray(events)).toBe(true);
  }, 40_000);

  test('respects maxResults parameter', async () => {
    const start = new Date();
    const end = new Date(Date.now() + 30 * 24 * 60 * 60 * 1000);
    const events = await client.getCalendarEvents(start.toISOString(), end.toISOString(), { maxResults: 3 });
    expect(events.length).toBeLessThanOrEqual(3);
  }, 40_000);
});
