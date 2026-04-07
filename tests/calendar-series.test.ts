// tests/calendar-series.test.ts
import { TokenManager } from '../src/auth.js';
import { CalendarClient } from '../src/calendar.js';
import type { OwaCreateEventPayload, CalendarEvent } from '../src/types.js';

describe('CalendarClient series operations', () => {
  let manager: TokenManager;
  let client: CalendarClient;

  beforeAll(async () => {
    manager = new TokenManager();
    client = new CalendarClient(manager);
  });

  afterAll(async () => {
    await manager.close();
  });

  describe('series metadata in calendarview', () => {
    test('events include type field', async () => {
      const start = new Date();
      const end = new Date(Date.now() + 14 * 24 * 60 * 60 * 1000);
      const events = await client.getCalendarEvents(start.toISOString(), end.toISOString());

      expect(events.length).toBeGreaterThan(0);
      for (const event of events) {
        expect(['singleInstance', 'occurrence', 'exception', 'seriesMaster']).toContain(event.type);
        expect(event).toHaveProperty('seriesMasterId');
        expect(event).toHaveProperty('recurrence');
      }
    }, 40_000);

    test('recurring occurrences have seriesMasterId set', async () => {
      const start = new Date();
      const end = new Date(Date.now() + 14 * 24 * 60 * 60 * 1000);
      const events = await client.getCalendarEvents(start.toISOString(), end.toISOString());

      const occurrences = events.filter(e => e.type === 'occurrence');
      if (occurrences.length === 0) {
        console.warn('No recurring occurrences found in date range — skipping seriesMasterId assertion');
        return;
      }
      for (const occ of occurrences) {
        expect(occ.seriesMasterId).toBeTruthy();
        expect(occ.isRecurring).toBe(true);
      }
    }, 40_000);

    test('single instances have null seriesMasterId and null recurrence', async () => {
      const start = new Date();
      const end = new Date(Date.now() + 14 * 24 * 60 * 60 * 1000);
      const events = await client.getCalendarEvents(start.toISOString(), end.toISOString());

      const singles = events.filter(e => e.type === 'singleInstance');
      if (singles.length === 0) {
        console.warn('No single-instance events found — skipping assertion');
        return;
      }
      for (const s of singles) {
        expect(s.seriesMasterId).toBeNull();
        expect(s.recurrence).toBeNull();
        expect(s.isRecurring).toBe(false);
      }
    }, 40_000);
  });

  describe('resolveSeriesMasterId', () => {
    test('resolves occurrence to series master ID', async () => {
      const start = new Date();
      const end = new Date(Date.now() + 14 * 24 * 60 * 60 * 1000);
      const events = await client.getCalendarEvents(start.toISOString(), end.toISOString());
      const occurrence = events.find(e => e.type === 'occurrence');
      if (!occurrence) {
        console.warn('No recurring occurrence found — skipping resolveSeriesMasterId test');
        return;
      }

      const masterId = await client.resolveSeriesMasterId(occurrence.id);
      expect(masterId).toBeTruthy();
      expect(masterId).toBe(occurrence.seriesMasterId);
    }, 40_000);

    test('returns own ID for series master', async () => {
      const start = new Date();
      const end = new Date(Date.now() + 14 * 24 * 60 * 60 * 1000);
      const events = await client.getCalendarEvents(start.toISOString(), end.toISOString());
      const occurrence = events.find(e => e.type === 'occurrence' && e.seriesMasterId);
      if (!occurrence) {
        console.warn('No recurring occurrence found — skipping series master resolution test');
        return;
      }

      const masterId = await client.resolveSeriesMasterId(occurrence.seriesMasterId!);
      expect(masterId).toBe(occurrence.seriesMasterId);
    }, 40_000);

    test('throws for single instance events', async () => {
      const start = new Date();
      const end = new Date(Date.now() + 14 * 24 * 60 * 60 * 1000);
      const events = await client.getCalendarEvents(start.toISOString(), end.toISOString());
      const single = events.find(e => e.type === 'singleInstance');
      if (!single) {
        console.warn('No single-instance event found — skipping test');
        return;
      }

      await expect(client.resolveSeriesMasterId(single.id))
        .rejects.toThrow('not part of a recurring series');
    }, 40_000);
  });

  describe('getSeriesMaster', () => {
    test('returns master with recurrence from occurrence ID', async () => {
      const start = new Date();
      const end = new Date(Date.now() + 14 * 24 * 60 * 60 * 1000);
      const events = await client.getCalendarEvents(start.toISOString(), end.toISOString());
      const occurrence = events.find(e => e.type === 'occurrence');
      if (!occurrence) {
        console.warn('No recurring occurrence found — skipping getSeriesMaster test');
        return;
      }

      const master = await client.getSeriesMaster(occurrence.id);
      expect(master.type).toBe('seriesmaster');
      expect(master.id).toBe(occurrence.seriesMasterId);
      expect(master.recurrence).not.toBeNull();
      expect(master.recurrence!.pattern).toHaveProperty('type');
      expect(master.recurrence!.pattern).toHaveProperty('interval');
      expect(master.recurrence!.range).toHaveProperty('type');
      expect(master.recurrence!.range).toHaveProperty('startDate');
      expect(master).toHaveProperty('cancelledOccurrences');
      expect(Array.isArray(master.cancelledOccurrences)).toBe(true);
    }, 40_000);
  });

  describe('listSeriesInstances', () => {
    test('returns instances for a recurring series', async () => {
      const start = new Date();
      const end = new Date(Date.now() + 14 * 24 * 60 * 60 * 1000);
      const events = await client.getCalendarEvents(start.toISOString(), end.toISOString());
      const occurrence = events.find(e => e.type === 'occurrence');
      if (!occurrence) {
        console.warn('No recurring occurrence found — skipping listSeriesInstances test');
        return;
      }

      // List instances over a 30-day window
      const rangeStart = new Date();
      const rangeEnd = new Date(Date.now() + 30 * 24 * 60 * 60 * 1000);
      const instances = await client.listSeriesInstances(
        occurrence.id,
        rangeStart.toISOString(),
        rangeEnd.toISOString()
      );

      expect(Array.isArray(instances)).toBe(true);
      expect(instances.length).toBeGreaterThan(0);
      for (const inst of instances) {
        expect(['occurrence', 'exception']).toContain(inst.type);
        expect(inst.seriesMasterId).toBe(occurrence.seriesMasterId);
      }
    }, 40_000);
  });

  describe('series cancel/delete operations', () => {
    let seriesEventId: string;

    afterEach(async () => {
      // Clean up any test series that wasn't already deleted
      if (seriesEventId) {
        try { await client.deleteEvent(seriesEventId, 'allInSeries'); } catch { /* ignore */ }
        seriesEventId = '';
      }
    });

    async function createTestSeries(): Promise<CalendarEvent> {
      const startDate = new Date(Date.now() + 7 * 24 * 60 * 60 * 1000);
      const startStr = startDate.toISOString().replace('Z', '').split('.')[0];
      const endDate = new Date(startDate.getTime() + 30 * 60 * 1000);
      const endStr = endDate.toISOString().replace('Z', '').split('.')[0];

      // Create a weekly recurring event for 5 weeks
      const payload: OwaCreateEventPayload = {
        Subject: `owa-mcp series test ${Date.now()}`,
        Start: { DateTime: startStr, TimeZone: 'UTC' },
        End: { DateTime: endStr, TimeZone: 'UTC' },
        ShowAs: 'Free',
        Sensitivity: 'Private',
        Recurrence: {
          Pattern: {
            Type: 'Weekly',
            Interval: 1,
            DaysOfWeek: [startDate.toLocaleDateString('en-US', { weekday: 'long' })],
            FirstDayOfWeek: 'Sunday',
          },
          Range: {
            Type: 'Numbered',
            StartDate: startStr.split('T')[0],
            NumberOfOccurrences: 5,
          },
        },
      };
      const event = await client.createEvent(payload);
      seriesEventId = event.id;
      return event;
    }

    test('cancel allInSeries removes the entire series', async () => {
      const created = await createTestSeries();

      // Cancel all
      await client.cancelEvent(created.id, 'Test cleanup', 'allInSeries');

      // Verify: listing instances over the next 60 days should return nothing or throw
      const rangeStart = new Date();
      const rangeEnd = new Date(Date.now() + 60 * 24 * 60 * 60 * 1000);

      try {
        const instances = await client.listSeriesInstances(
          created.id,
          rangeStart.toISOString(),
          rangeEnd.toISOString()
        );
        // If we get results, they should be empty (series was cancelled)
        expect(instances.length).toBe(0);
      } catch (err) {
        // API may return 404 for cancelled series — that's also acceptable
        expect(String(err)).toMatch(/404|not found|ErrorItemNotFound/i);
      }
      seriesEventId = ''; // Already cleaned up via cancel
    }, 60_000);

    test('delete allInSeries removes the entire series', async () => {
      const created = await createTestSeries();

      await client.deleteEvent(created.id, 'allInSeries');

      const rangeStart = new Date();
      const rangeEnd = new Date(Date.now() + 60 * 24 * 60 * 60 * 1000);

      try {
        const instances = await client.listSeriesInstances(
          created.id,
          rangeStart.toISOString(),
          rangeEnd.toISOString()
        );
        expect(instances.length).toBe(0);
      } catch (err) {
        expect(String(err)).toMatch(/404|not found|ErrorItemNotFound/i);
      }
      seriesEventId = ''; // Already cleaned up
    }, 60_000);

    test('thisAndFollowing truncates the series', async () => {
      const created = await createTestSeries();

      // List instances to find the third occurrence
      const rangeStart = new Date();
      const rangeEnd = new Date(Date.now() + 60 * 24 * 60 * 60 * 1000);
      const instances = await client.listSeriesInstances(
        created.id,
        rangeStart.toISOString(),
        rangeEnd.toISOString()
      );
      expect(instances.length).toBe(5);

      // Cancel from the 3rd occurrence onward
      const thirdInstance = instances[2];
      await client.cancelEvent(thirdInstance.id, undefined, 'thisAndFollowing');

      // Re-list: should now have only 2 instances
      const afterInstances = await client.listSeriesInstances(
        created.id,
        rangeStart.toISOString(),
        rangeEnd.toISOString()
      );
      expect(afterInstances.length).toBe(2);
    }, 60_000);
  });
});
