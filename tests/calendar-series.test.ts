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
});
