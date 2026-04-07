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

// ── Mail write types ────────────────────────────────────────

export interface OwaRecipient {
  EmailAddress: { Address: string; Name?: string };
}

export interface OwaSendMailPayload {
  Message: {
    Subject: string;
    Body: { ContentType: 'HTML' | 'Text'; Content: string };
    ToRecipients: OwaRecipient[];
    CcRecipients?: OwaRecipient[];
    BccRecipients?: OwaRecipient[];
    Importance?: 'Low' | 'Normal' | 'High';
  };
  SaveToSentItems?: boolean;
}

export interface OwaUpdateMailPayload {
  Subject?: string;
  Body?: { ContentType: 'HTML' | 'Text'; Content: string };
  ToRecipients?: OwaRecipient[];
  CcRecipients?: OwaRecipient[];
  BccRecipients?: OwaRecipient[];
  Importance?: 'Low' | 'Normal' | 'High';
  IsRead?: boolean;
  Flag?: { FlagStatus: 'NotFlagged' | 'Flagged' | 'Complete' };
}

// ── Mail types ──────────────────────────────────────────────

export interface MailFolder {
  id: string;
  displayName: string;
  parentFolderId: string;
  unreadCount: number;
  totalCount: number;
  childFolderCount: number;
}

export interface MailMessage {
  id: string;
  subject: string;
  from: string;
  toRecipients: string[];
  ccRecipients: string[];
  receivedDateTime: string;
  isRead: boolean;
  hasAttachments: boolean;
  importance: string;
  bodyPreview: string;
  body?: string;
  bodyType?: string;
  conversationId: string;
  flag: string;
  parentFolderId: string;
  attachments?: MailAttachment[];
}

export interface MailAttachment {
  id: string;
  name: string;
  contentType: string;
  size: number;
}

export interface MailAttachmentDownload {
  id: string;
  name: string;
  contentType: string;
  size: number;
  filePath: string;
}

// Raw OWA response shapes

export interface OwaMailFolder {
  Id: string;
  DisplayName: string;
  ParentFolderId: string;
  UnreadItemCount: number;
  TotalItemCount: number;
  ChildFolderCount: number;
}

export interface OwaMailMessage {
  Id: string;
  Subject: string;
  From: { EmailAddress: { Name: string; Address: string } };
  ToRecipients: { EmailAddress: { Name: string; Address: string } }[];
  CcRecipients: { EmailAddress: { Name: string; Address: string } }[];
  ReceivedDateTime: string;
  IsRead: boolean;
  HasAttachments: boolean;
  Importance: string;
  BodyPreview: string;
  Body?: { ContentType: string; Content: string };
  ConversationId: string;
  Flag: { FlagStatus: string };
  ParentFolderId: string;
  Attachments?: OwaMailAttachment[];
}

export interface OwaMailAttachment {
  Id: string;
  Name: string;
  ContentType: string;
  Size: number;
  ContentBytes?: string;
}

export interface OwaMailListResponse {
  value: OwaMailMessage[];
  '@odata.nextLink'?: string;
}

export interface OwaMailFolderListResponse {
  value: OwaMailFolder[];
}
