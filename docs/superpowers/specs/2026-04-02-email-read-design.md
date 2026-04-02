# Email Read Support — Design Spec

## Overview

Add email read capabilities to owa-mcp by creating a new `MailClient` class in `src/mail.ts` that follows the same pattern as `CalendarClient`. Five new MCP tools expose mail folder listing, message browsing, full-text search, single message reading, and attachment downloading.

No changes to `auth.ts` — the intercepted OWA token already carries `Mail.ReadWrite` scope.

## File Changes

### New files

| File | Purpose |
|------|---------|
| `src/mail.ts` | `MailClient` class — all mail read methods |
| `tests/mail.test.ts` | Integration tests for mail read tools |

### Modified files

| File | Change |
|------|--------|
| `src/types.ts` | Add mail-related types (normalised + raw OWA shapes) |
| `src/index.ts` | Instantiate `MailClient`, register 5 new tools |

### Untouched files

`src/auth.ts`, `src/calendar.ts` — no changes.

## Types (`src/types.ts`)

### Normalised output types

```typescript
interface MailFolder {
  id: string;
  displayName: string;
  parentFolderId: string;
  unreadCount: number;
  totalCount: number;
  childFolderCount: number;
}

interface MailMessage {
  id: string;
  subject: string;
  from: string;             // "Name <email>" or just "email"
  toRecipients: string[];   // same format
  ccRecipients: string[];
  receivedDateTime: string; // ISO 8601
  isRead: boolean;
  hasAttachments: boolean;
  importance: string;       // Low | Normal | High
  bodyPreview: string;
  body?: string;            // only populated by get_email
  bodyType?: string;        // "text" | "html", only when body is present
  conversationId: string;
  flag: string;             // NotFlagged | Flagged | Complete
  parentFolderId: string;
  attachments?: MailAttachment[];  // only populated by get_email
}

interface MailAttachment {
  id: string;
  name: string;
  contentType: string;
  size: number;             // bytes
}

interface MailAttachmentDownload {
  id: string;
  name: string;
  contentType: string;
  size: number;
  filePath: string;         // absolute path to downloaded file
}
```

### Raw OWA response types

```typescript
interface OwaMailFolder {
  Id: string;
  DisplayName: string;
  ParentFolderId: string;
  UnreadItemCount: number;
  TotalItemCount: number;
  ChildFolderCount: number;
}

interface OwaMailMessage {
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

interface OwaMailAttachment {
  Id: string;
  Name: string;
  ContentType: string;
  Size: number;
  ContentBytes?: string;    // base64, only from attachment download
}

interface OwaMailListResponse {
  value: OwaMailMessage[];
  '@odata.nextLink'?: string;
}

interface OwaMailFolderListResponse {
  value: OwaMailFolder[];
}
```

## Tools

### `list_mail_folders`

List mail folders in the mailbox.

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `parentFolderId` | string | no | (root) | List children of a specific folder |

- API: `GET /me/mailfolders` or `GET /me/mailfolders/{id}/childfolders`
- `$select`: Id, DisplayName, ParentFolderId, UnreadItemCount, TotalItemCount, ChildFolderCount
- Returns: `MailFolder[]`

### `get_emails`

List messages in a folder with optional preset filters.

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `folderId` | string | no | Inbox | Folder ID or well-known name |
| `filter` | enum | no | all | "all", "unread", "flagged", "today", "this_week" |
| `limit` | number | no | 20 | Max messages to return (max 500) |
| `pageToken` | string | no | — | Pagination token from previous response |

- API: `GET /me/mailfolders/{folderId}/messages`
- `$select`: Id, Subject, From, ToRecipients, CcRecipients, ReceivedDateTime, IsRead, HasAttachments, Importance, BodyPreview, ConversationId, Flag, ParentFolderId
- `$orderby`: ReceivedDateTime desc
- `$top`: limit
- `$skip`: parsed from pageToken (integer offset)
- `$filter` presets:
  - `unread` → `IsRead eq false`
  - `flagged` → `Flag/FlagStatus eq 'Flagged'`
  - `today` → `ReceivedDateTime ge {start of today in UTC}`
  - `this_week` → `ReceivedDateTime ge {7 days ago in UTC}`
- Body content is NOT included (only `bodyPreview`)
- Returns: `{ messages: MailMessage[], nextPageToken?: string }`

### `search_emails`

Search messages using full-text search OR structured filters (mutually exclusive).

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `query` | string | no | — | Full-text search (uses `$search`) |
| `from` | string | no | — | Filter by sender email/name |
| `subject` | string | no | — | Filter by subject contains |
| `hasAttachments` | boolean | no | — | Filter by attachment presence |
| `receivedAfter` | string | no | — | ISO 8601 datetime lower bound |
| `receivedBefore` | string | no | — | ISO 8601 datetime upper bound |
| `folderId` | string | no | — | Scope search to a folder |
| `limit` | number | no | 20 | Max results (max 500) |

- If `query` is provided: `GET /me/messages?$search="query"` (or scoped to folder). Structured filter params must NOT be provided — throw error if both `query` and any filter param are set.
- If no `query`: build `$filter` from structured params. Combine with `and`.
  - `from` → `From/EmailAddress/Address eq 'value'`
  - `subject` → `contains(Subject, 'value')`
  - `hasAttachments` → `HasAttachments eq true/false`
  - `receivedAfter` → `ReceivedDateTime ge value`
  - `receivedBefore` → `ReceivedDateTime le value`
- Same `$select` and normalisation as `get_emails`
- Returns: `{ messages: MailMessage[], nextPageToken?: string }`

### `get_email`

Read a single message with full body content.

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `messageId` | string | yes | — | Message ID |
| `format` | enum | no | text | "text" or "html" |

- API: `GET /me/messages/{id}`
- `$select`: all fields including Body, plus `$expand=Attachments($select=Id,Name,ContentType,Size)` to include attachment metadata
- `Prefer: outlook.body-content-type="text"` or `"html"` based on `format` param
- Returns: `MailMessage` with `body`, `bodyType`, and `attachments` populated

### `get_attachment`

Download an attachment to disk.

| Parameter | Type | Required | Default | Description |
|-----------|------|----------|---------|-------------|
| `messageId` | string | yes | — | Message ID |
| `attachmentId` | string | yes | — | Attachment ID from get_email |

- API: `GET /me/messages/{messageId}/attachments/{attachmentId}`
- Writes `ContentBytes` (base64-decoded) to `os.tmpdir()/owa-mcp-attachments/{sanitised-filename}`
- Filename sanitisation: replace path separators and control chars, truncate to 200 chars, overwrite if file exists
- Creates the output directory if it doesn't exist
- Returns: `MailAttachmentDownload` with absolute `filePath`

## `MailClient` Class (`src/mail.ts`)

### Constructor

```typescript
constructor(private readonly tokens: TokenManager) {}
```

### Private methods

- `request(method, path, options?)` — same implementation as `CalendarClient.request()`. ~20 lines of duplication, acceptable. Handles auth header, timezone preference, JSON body, error throwing.
- `normaliseMessage(raw: OwaMailMessage): MailMessage` — maps PascalCase to camelCase. Formats `from` as `"Name <email>"`. Omits `body` unless present in raw response.
- `normaliseFolder(raw: OwaMailFolder): MailFolder` — maps PascalCase to camelCase.
- `formatRecipient(r: { EmailAddress: { Name: string; Address: string } }): string` — returns `"Name <email>"` or just `"email"` if name is empty/matches address.

### Public methods

- `listFolders(parentFolderId?: string): Promise<MailFolder[]>`
- `getMessages(folderId: string, options: GetMessagesOptions): Promise<{ messages: MailMessage[]; nextPageToken?: string }>`
- `searchMessages(options: SearchMessagesOptions): Promise<{ messages: MailMessage[]; nextPageToken?: string }>`
- `getMessage(messageId: string, format: 'text' | 'html'): Promise<MailMessage>`
- `getAttachment(messageId: string, attachmentId: string): Promise<MailAttachmentDownload>`

### Pagination

Uses `$skip`/`$top`. The `nextPageToken` returned to the LLM is the string representation of the next skip offset. If the API returns fewer results than `$top`, no `nextPageToken` is returned.

## Tool Registration (`src/index.ts`)

- Instantiate `const mailClient = new MailClient(tokenManager);` alongside `calendarClient`
- Register 5 tools using `server.tool()` with Zod schemas
- Handlers call `MailClient` methods, return `JSON.stringify(result, null, 2)` in `{ content: [{ type: 'text', text }] }`
- No pre-formatting or field truncation

## Testing (`tests/mail.test.ts`)

Integration tests (same as calendar tests — require live Edge session):

1. `list_mail_folders` — returns array with at least Inbox
2. `get_emails` — returns messages from Inbox, verifies shape
3. `get_emails` with filter — `unread` filter returns only unread messages
4. `search_emails` with query — full-text search returns results
5. `search_emails` with structured filters — filter by `from` returns results
6. `search_emails` — error when both `query` and structured filters provided
7. `get_email` — returns full message with body in text format
8. `get_email` with html format — returns HTML body
9. `get_attachment` — downloads attachment, file exists at returned path
