// src/calendar.ts
import type { TokenManager } from './auth.js';
import type { CalendarEvent, OwaCalendarViewResponse, OwaCalendarEvent } from './types.js';

const OWA_BASE = 'https://outlook.office.com/api/v2.0';

export interface GetCalendarEventsOptions {
  maxResults?: number;          // default 50
  timezone?: string;            // IANA tz name, default 'UTC'
}

export class CalendarClient {
  constructor(private readonly tokens: TokenManager) {}

  /**
   * Returns calendar events between startDateTime and endDateTime (ISO 8601 strings).
   * Handles OData paging automatically up to maxResults.
   */
  async getCalendarEvents(
    startDateTime: string,
    endDateTime: string,
    options: GetCalendarEventsOptions = {}
  ): Promise<CalendarEvent[]> {
    const { maxResults = 50, timezone = 'UTC' } = options;

    const token = await this.tokens.getToken();
    const params = new URLSearchParams({
      startDateTime,
      endDateTime,
      '$select': 'Id,Subject,Start,End,IsAllDay,Organizer,Location,IsOnlineMeeting,ShowAs,Recurrence,Sensitivity,BodyPreview',
      '$top': String(Math.min(maxResults, 100)),
      '$orderby': 'Start/DateTime asc',
    });

    const url = `${OWA_BASE}/me/calendarview?${params}`;
    const events: CalendarEvent[] = [];
    let nextLink: string | undefined = url;

    while (nextLink && events.length < maxResults) {
      const res = await fetch(nextLink, {
        headers: {
          Authorization: `Bearer ${token.value}`,
          Accept: 'application/json',
          Prefer: `outlook.timezone="${timezone}"`,
        },
      });

      if (!res.ok) {
        const body = await res.text();
        throw new Error(`OWA calendar API error ${res.status}: ${body}`);
      }

      const data = (await res.json()) as OwaCalendarViewResponse;
      for (const raw of data.value) {
        events.push(this.normalise(raw));
        if (events.length >= maxResults) break;
      }
      nextLink = data['@odata.nextLink'];
    }

    return events;
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
