// tests/calendar-write.test.ts
import { TokenManager } from '../src/auth.js';
import { CalendarClient } from '../src/calendar.js';
import type { OwaCreateEventPayload } from '../src/types.js';

describe('CalendarClient write operations', () => {
  let manager: TokenManager;
  let client: CalendarClient;
  let createdEventId: string;

  beforeAll(async () => {
    manager = new TokenManager();
    client = new CalendarClient(manager);
  });

  afterAll(async () => {
    if (createdEventId) {
      try { await client.deleteEvent(createdEventId); } catch { /* ignore */ }
    }
    await manager.close();
  });

  test('creates an event', async () => {
    const futureDate = new Date(Date.now() + 7 * 24 * 60 * 60 * 1000);
    const startStr = futureDate.toISOString().replace('Z', '').split('.')[0];
    const endDate = new Date(futureDate.getTime() + 30 * 60 * 1000);
    const endStr = endDate.toISOString().replace('Z', '').split('.')[0];

    const payload: OwaCreateEventPayload = {
      Subject: `owa-mcp test event ${Date.now()}`,
      Start: { DateTime: startStr, TimeZone: 'UTC' },
      End: { DateTime: endStr, TimeZone: 'UTC' },
      ShowAs: 'Free',
      Sensitivity: 'Private',
    };
    const event = await client.createEvent(payload);
    expect(event.id).toBeTruthy();
    expect(event.subject).toMatch(/owa-mcp test event/);
    expect(event.showAs).toBe('Free');
    expect(event.isPrivate).toBe(true);
    createdEventId = event.id;
  }, 40_000);

  test('updates the event', async () => {
    const event = await client.updateEvent(createdEventId, {
      Subject: 'owa-mcp updated test event',
      ShowAs: 'Busy',
    });
    expect(event.subject).toBe('owa-mcp updated test event');
    expect(event.showAs).toBe('Busy');
  }, 20_000);

  test('RSVPs to the event via service.svc with REST fallback', async () => {
    // respondToEvent tries service.svc first (bypasses ResponseRequested: false),
    // then falls back to REST API if service.svc can't resolve the event ID.
    await client.respondToEvent(createdEventId, 'tentativelyaccept', {
      SendResponse: false,
    });

    // Verify the response was recorded
    const events = await client.getCalendarEvents(
      new Date(Date.now() + 6 * 24 * 60 * 60 * 1000).toISOString(),
      new Date(Date.now() + 8 * 24 * 60 * 60 * 1000).toISOString(),
    );
    const found = events.find(e => e.id === createdEventId);
    expect(found).toBeTruthy();
    expect(found!.showAs).toBe('Tentative');

    // Accept it back
    await client.respondToEvent(createdEventId, 'accept', {
      SendResponse: false,
    });
  }, 40_000);

  test('deletes the event', async () => {
    await client.deleteEvent(createdEventId);
    createdEventId = '';
  }, 20_000);
});
