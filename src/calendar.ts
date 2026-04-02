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
      '$select': 'Id,Subject,Start,End,IsAllDay,Organizer,Location,IsOnlineMeeting,ShowAs,Recurrence,Sensitivity,BodyPreview',
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
    payload: OwaRsvpPayload = {}
  ): Promise<void> {
    await this.request('POST', `/me/events/${eventId}/${action}`, { body: payload });
  }

  /**
   * Follow a calendar event using OWA's native Follow protocol.
   * Sends a "Following" notification to the organizer. The event appears
   * on the user's calendar with ShowAs=Free and subject prefixed "Following:".
   * Uses the OWA service.svc internal API with Attendance=3, Mode=3.
   */
  async followEvent(eventId: string, comment?: string, timezone?: string): Promise<CalendarEvent> {
    const token = await this.tokens.getToken();
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
        EventId: { __type: 'ItemId:#Exchange', Id: eventId },
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
      throw new Error(`OWA Follow error: ${svcBody.Body.ResponseCode} — ${svcBody.Body.MessageText ?? ''}`);
    }

    // Fetch the updated event to return its new state
    const updated = await this.request('GET', `/me/events/${eventId}`, { timezone });
    const raw = (await updated.json()) as OwaCalendarEvent;
    return this.normalise(raw);
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
