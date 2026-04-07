// src/calendar.ts
import type { TokenManager } from './auth.js';
import type {
  CalendarEvent,
  OwaCalendarViewResponse,
  OwaCalendarEvent,
  OwaCreateEventPayload,
  OwaUpdateEventPayload,
  RsvpAction,
  OwaRsvpPayload,
  RecurrenceInfo,
} from './types.js';

const OWA_BASE = 'https://outlook.office.com/api/v2.0';
const OWA_SVC = 'https://outlook.office.com/owa/service.svc';

export interface GetCalendarEventsOptions {
  maxResults?: number;          // default 50
  timezone?: string;            // IANA tz name, default 'UTC'
}

export class CalendarClient {
  constructor(private readonly tokens: TokenManager) {}

  async getCalendarEvents(
    startDateTime: string,
    endDateTime: string,
    options: GetCalendarEventsOptions = {}
  ): Promise<CalendarEvent[]> {
    const { maxResults = 50, timezone = 'UTC' } = options;

    const params = new URLSearchParams({
      startDateTime,
      endDateTime,
      '$select': 'Id,Subject,Start,End,IsAllDay,Organizer,Location,IsOnlineMeeting,ShowAs,Recurrence,Sensitivity,BodyPreview,Type,SeriesMasterId',
      '$top': String(Math.min(maxResults, 100)),
      '$orderby': 'Start/DateTime asc',
    });

    const events: CalendarEvent[] = [];
    let nextLink: string | undefined = `${OWA_BASE}/me/calendarview?${params}`;

    while (nextLink && events.length < maxResults) {
      const res = await this.request('GET', nextLink, { timezone });
      const data = (await res.json()) as OwaCalendarViewResponse;
      for (const raw of data.value) {
        events.push(this.normalise(raw));
        if (events.length >= maxResults) break;
      }
      nextLink = data['@odata.nextLink'];
    }

    return events;
  }

  async createEvent(
    payload: OwaCreateEventPayload,
    timezone?: string
  ): Promise<CalendarEvent> {
    const res = await this.request('POST', '/me/events', { body: payload, timezone });
    const raw = (await res.json()) as OwaCalendarEvent;
    return this.normalise(raw);
  }

  async updateEvent(
    eventId: string,
    payload: OwaUpdateEventPayload,
    timezone?: string
  ): Promise<CalendarEvent> {
    const res = await this.request('PATCH', `/me/events/${eventId}`, { body: payload, timezone });
    const raw = (await res.json()) as OwaCalendarEvent;
    return this.normalise(raw);
  }

  async cancelEvent(eventId: string, comment?: string): Promise<void> {
    await this.request('POST', `/me/events/${eventId}/cancel`, {
      body: comment ? { Comment: comment } : {},
    });
  }

  async deleteEvent(eventId: string): Promise<void> {
    await this.request('DELETE', `/me/events/${eventId}`);
  }

  async respondToEvent(
    eventId: string,
    action: RsvpAction,
    payload: OwaRsvpPayload = {},
    timezone?: string
  ): Promise<void> {
    // Try service.svc first — bypasses ResponseRequested: false restriction.
    // Falls back to REST API if service.svc can't resolve the event ID.
    const token = await this.tokens.getToken();
    try {
      const svcEventId = await this.toServiceId(eventId, token.value);
      const responseMap: Record<RsvpAction, string> = {
        accept: 'Accept',
        tentativelyaccept: 'Tentative',
        decline: 'Decline',
      };

      const svcPayload = {
        __type: 'RespondToCalendarEventJsonRequest:#Exchange',
        Header: {
          __type: 'JsonRequestHeaders:#Exchange',
          RequestServerVersion: 'V2018_01_08',
          TimeZoneContext: {
            __type: 'TimeZoneContext:#Exchange',
            TimeZoneDefinition: {
              __type: 'TimeZoneDefinitionType:#Exchange',
              Id: timezone ?? 'W. Europe Standard Time',
            },
          },
        },
        Body: {
          __type: 'RespondToCalendarEventRequest:#Exchange',
          EventId: { __type: 'ItemId:#Exchange', Id: svcEventId },
          Response: responseMap[action],
          SendResponse: payload.SendResponse ?? true,
          Notes: payload.Comment
            ? { __type: 'BodyContentType:#Exchange', BodyType: 'HTML', Value: `<div>${payload.Comment}</div>` }
            : { __type: 'BodyContentType:#Exchange', BodyType: 'HTML', Value: '<div><br></div>' },
          ProposedStartTime: payload.ProposedNewTime?.Start?.DateTime ?? '',
          ProposedEndTime: payload.ProposedNewTime?.End?.DateTime ?? '',
          Attendance: 0,
          Mode: 0,
        },
      };

      const res = await fetch(`${OWA_SVC}?action=RespondToCalendarEvent`, {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${token.value}`,
          'Content-Type': 'application/json; charset=utf-8',
          action: 'RespondToCalendarEvent',
          'x-owa-urlpostdata': encodeURIComponent(JSON.stringify(svcPayload)),
          'x-req-source': 'Calendar',
        },
      });

      const svcBody = (await res.json()) as { Body: { ResponseCode: string; MessageText?: string } };
      if (svcBody.Body.ResponseCode === 'NoError') {
        return; // service.svc succeeded
      }
      // Fall through to REST API
    } catch {
      // Fall through to REST API
    }

    // Fallback: REST API (subject to ResponseRequested restriction)
    await this.request('POST', `/me/events/${eventId}/${action}`, { body: payload });
  }

  /**
   * Follow a calendar event using OWA's native Follow protocol.
   * Sends a "Following" notification to the organizer. The event appears
   * on the user's calendar with ShowAs=Free and subject prefixed "Following:".
   * Uses the OWA service.svc internal API with Attendance=3, Mode=3.
   * Falls back to tentativelyAccept + PATCH for events where service.svc fails.
   */
  async followEvent(eventId: string, comment?: string, timezone?: string): Promise<CalendarEvent> {
    const token = await this.tokens.getToken();
    const svcEventId = await this.toServiceId(eventId, token.value);
    const payload = {
      __type: 'RespondToCalendarEventJsonRequest:#Exchange',
      Header: {
        __type: 'JsonRequestHeaders:#Exchange',
        RequestServerVersion: 'V2018_01_08',
        TimeZoneContext: {
          __type: 'TimeZoneContext:#Exchange',
          TimeZoneDefinition: {
            __type: 'TimeZoneDefinitionType:#Exchange',
            Id: timezone ?? 'W. Europe Standard Time',
          },
        },
      },
      Body: {
        __type: 'RespondToCalendarEventRequest:#Exchange',
        EventId: { __type: 'ItemId:#Exchange', Id: svcEventId },
        Response: 'Tentative',
        SendResponse: true,
        Notes: {
          __type: 'BodyContentType:#Exchange',
          BodyType: 'HTML',
          Value: comment ? `<div>${comment}</div>` : '<div><br></div>',
        },
        ProposedStartTime: '',
        ProposedEndTime: '',
        Attendance: 3,
        Mode: 3,
      },
    };

    const res = await fetch(`${OWA_SVC}?action=RespondToCalendarEvent`, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token.value}`,
        'Content-Type': 'application/json; charset=utf-8',
        action: 'RespondToCalendarEvent',
        'x-owa-urlpostdata': encodeURIComponent(JSON.stringify(payload)),
        'x-req-source': 'Calendar',
      },
    });

    const svcBody = (await res.json()) as { Body: { ResponseCode: string; MessageText?: string } };
    if (svcBody.Body.ResponseCode !== 'NoError') {
      // Fallback: tentativelyAccept + PATCH ShowAs=Free
      // This handles single-instance events where service.svc can't resolve the ID.
      // Note: organizer is notified only if a comment is provided.
      const sendResponse = !!comment;
      await this.request('POST', `/me/events/${eventId}/tentativelyaccept`, {
        body: { Comment: comment ?? '', SendResponse: sendResponse },
      });
      await this.request('PATCH', `/me/events/${eventId}`, {
        body: { ShowAs: 'Free' },
      });
    }

    const updated = await this.request('GET', `/me/events/${eventId}`, { timezone });
    const raw = (await updated.json()) as OwaCalendarEvent;
    return this.normalise(raw);
  }

  /**
   * Resolve any event ID to its series master ID.
   * Returns the event's own ID if it is already a series master.
   * Throws if the event is a singleInstance (not part of a series).
   */
  async resolveSeriesMasterId(eventId: string): Promise<string> {
    const res = await this.request('GET', `/me/events/${eventId}?$select=Id,Type,SeriesMasterId`);
    const raw = (await res.json()) as { Id: string; Type: string; SeriesMasterId: string | null };
    const type = raw.Type?.toLowerCase();

    if (type === 'singleinstance') {
      throw new Error('This event is not part of a recurring series');
    }
    if (type === 'seriesmaster') {
      return raw.Id;
    }
    // occurrence or exception
    if (!raw.SeriesMasterId) {
      throw new Error(`Event ${eventId} is ${type} but has no SeriesMasterId`);
    }
    return raw.SeriesMasterId;
  }

  /**
   * Fetch the series master event and return its current recurrence pattern.
   * Used by thisAndFollowing scope to construct the PATCH payload.
   */
  async getSeriesRecurrence(seriesMasterId: string): Promise<{ recurrence: unknown; event: OwaCalendarEvent }> {
    const res = await this.request('GET', `/me/events/${seriesMasterId}?$select=Id,Subject,Start,End,Recurrence,Type,IsAllDay,Organizer,Location,IsOnlineMeeting,ShowAs,Sensitivity,BodyPreview,SeriesMasterId`);
    const raw = (await res.json()) as OwaCalendarEvent & { Recurrence: unknown };
    if (!raw.Recurrence) {
      throw new Error(`Series master ${seriesMasterId} has no Recurrence`);
    }
    return { recurrence: raw.Recurrence, event: raw };
  }

  /** Translate a REST API event ID to the format service.svc expects. */
  private async toServiceId(restId: string, token: string): Promise<string> {
    const res = await fetch(`${OWA_BASE.replace('/v2.0', '/beta')}/me/translateExchangeIds`, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        InputIds: [restId],
        TargetIdType: 'RestImmutableEntryId',
        SourceIdType: 'RestId',
      }),
    });
    if (!res.ok) {
      throw new Error(`ID translation failed: ${res.status}`);
    }
    const data = (await res.json()) as { value: { TargetId: string }[] };
    // Convert base64url to standard base64 for service.svc
    return data.value[0].TargetId.replace(/-/g, '+').replace(/_/g, '/');
  }

  private async request(
    method: string,
    path: string,
    options: { body?: unknown; timezone?: string } = {}
  ): Promise<Response> {
    const token = await this.tokens.getToken();
    const headers: Record<string, string> = {
      Authorization: `Bearer ${token.value}`,
      Accept: 'application/json',
    };
    if (options.timezone) {
      headers['Prefer'] = `outlook.timezone="${options.timezone}"`;
    }
    if (options.body !== undefined) {
      headers['Content-Type'] = 'application/json';
    }

    const url = path.startsWith('http') ? path : `${OWA_BASE}${path}`;
    const res = await fetch(url, {
      method,
      headers,
      body: options.body !== undefined ? JSON.stringify(options.body) : undefined,
    });

    if (!res.ok) {
      const text = await res.text();
      throw new Error(`OWA API error ${res.status} ${method} ${path}: ${text}`);
    }
    return res;
  }

  private normalise(raw: OwaCalendarEvent): CalendarEvent {
    const type = (raw.Type?.toLowerCase() ?? 'singleInstance') as CalendarEvent['type'];
    const recurrence = raw.Recurrence
      ? this.normaliseRecurrence(raw.Recurrence)
      : null;

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
      isRecurring: type !== 'singleInstance',
      isPrivate: raw.Sensitivity === 'Private',
      bodyPreview: raw.BodyPreview ?? '',
      type,
      seriesMasterId: raw.SeriesMasterId ?? null,
      recurrence,
    };
  }

  private normaliseRecurrence(raw: unknown): RecurrenceInfo | null {
    if (!raw || typeof raw !== 'object') return null;
    const rec = raw as Record<string, unknown>;
    const pattern = rec.Pattern as Record<string, unknown> | undefined;
    const range = rec.Range as Record<string, unknown> | undefined;
    if (!pattern || !range) return null;

    return {
      pattern: {
        type: String(pattern.Type ?? '').toLowerCase(),
        interval: Number(pattern.Interval ?? 1),
        daysOfWeek: Array.isArray(pattern.DaysOfWeek)
          ? pattern.DaysOfWeek.map((d: unknown) => String(d).toLowerCase())
          : undefined,
        dayOfMonth: pattern.DayOfMonth != null ? Number(pattern.DayOfMonth) : undefined,
        month: pattern.Month != null ? Number(pattern.Month) : undefined,
        index: pattern.Index != null ? String(pattern.Index).toLowerCase() : undefined,
        firstDayOfWeek: pattern.FirstDayOfWeek != null
          ? String(pattern.FirstDayOfWeek).toLowerCase()
          : undefined,
      },
      range: {
        type: String(range.Type ?? '').toLowerCase(),
        startDate: String(range.StartDate ?? ''),
        endDate: range.EndDate != null ? String(range.EndDate) : undefined,
        numberOfOccurrences: range.NumberOfOccurrences != null
          ? Number(range.NumberOfOccurrences)
          : undefined,
        recurrenceTimeZone: range.RecurrenceTimeZone != null
          ? String(range.RecurrenceTimeZone)
          : undefined,
      },
    };
  }
}
