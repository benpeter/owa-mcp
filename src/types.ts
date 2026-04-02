// src/types.ts
// tva

export interface OwaToken {
  value: string;       // raw JWT
  expiresAt: number;   // unix epoch ms
  issuedAt: number;    // unix epoch ms
}

export interface CalendarEvent {
  id: string;
  subject: string;
  start: string;       // ISO 8601
  end: string;         // ISO 8601
  isAllDay: boolean;
  organizer: string;
  location: string;
  isOnlineMeeting: boolean;
  showAs: string;      // Free | Tentative | Busy | Oof | WorkingElsewhere | Unknown
  isRecurring: boolean;
  isPrivate: boolean;
  bodyPreview: string;
}

export interface OwaCalendarViewResponse {
  value: OwaCalendarEvent[];
  '@odata.nextLink'?: string;
}

// Raw shape returned by outlook.office.com/api/v2.0/me/calendarview
export interface OwaCalendarEvent {
  Id: string;
  Subject: string;
  Start: { DateTime: string; TimeZone: string };
  End: { DateTime: string; TimeZone: string };
  IsAllDay: boolean;
  Organizer: { EmailAddress: { Name: string; Address: string } };
  Location: { DisplayName: string };
  IsOnlineMeeting: boolean;
  ShowAs: string;
  IsReminderOn: boolean;
  Recurrence: unknown | null;
  Sensitivity: string;   // Normal | Personal | Private | Confidential
  BodyPreview: string;
}

// Request payload for POST /me/events (PascalCase for OWA API)
export interface OwaCreateEventPayload {
  Subject: string;
  Start: { DateTime: string; TimeZone: string };
  End: { DateTime: string; TimeZone: string };
  Body?: { ContentType: 'HTML' | 'Text'; Content: string };
  Location?: { DisplayName: string };
  Attendees?: OwaAttendee[];
  IsAllDay?: boolean;
  ShowAs?: string;
  Importance?: 'Low' | 'Normal' | 'High';
  Sensitivity?: 'Normal' | 'Personal' | 'Private' | 'Confidential';
  IsReminderOn?: boolean;
  ReminderMinutesBeforeStart?: number;
  IsOnlineMeeting?: boolean;
  Recurrence?: unknown;
}

export interface OwaAttendee {
  EmailAddress: { Address: string; Name?: string };
  Type: 'Required' | 'Optional' | 'Resource';
}

// Request payload for PATCH /me/events/{id}
export type OwaUpdateEventPayload = Partial<OwaCreateEventPayload>;

// RSVP action types
export type RsvpAction = 'accept' | 'tentativelyaccept' | 'decline';

export interface OwaRsvpPayload {
  Comment?: string;
  SendResponse?: boolean;
  ProposedNewTime?: {
    Start: { DateTime: string; TimeZone: string };
    End: { DateTime: string; TimeZone: string };
  };
}
