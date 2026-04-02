# Email Read Support Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add 5 MCP tools for reading email — folder listing, message browsing, search, single message reading, and attachment downloading.

**Architecture:** New `MailClient` class in `src/mail.ts` mirrors the `CalendarClient` pattern. Types added to `src/types.ts`. Tools registered in `src/index.ts`. No changes to auth or calendar code.

**Tech Stack:** TypeScript, OWA REST API v2.0, `@modelcontextprotocol/sdk`, Zod, Jest

---

### Task 1: Add mail types to `src/types.ts`

**Files:**
- Modify: `src/types.ts`

- [ ] **Step 1: Add all mail types**

Append the following to the end of `src/types.ts`:

```typescript
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
```

- [ ] **Step 2: Verify it compiles**

Run: `npx tsc --noEmit`
Expected: no errors

- [ ] **Step 3: Commit**

```bash
git add src/types.ts
git commit -m "feat: add mail-related types for email read support"
```

---

### Task 2: Create `MailClient` — `listFolders` method

**Files:**
- Create: `src/mail.ts`
- Create: `tests/mail.test.ts`

- [ ] **Step 1: Write the test file skeleton and `listFolders` test**

Create `tests/mail.test.ts`:

```typescript
import { TokenManager } from '../src/auth.js';
import { MailClient } from '../src/mail.js';

describe('MailClient', () => {
  let manager: TokenManager;
  let client: MailClient;

  beforeAll(async () => {
    manager = new TokenManager();
    client = new MailClient(manager);
  });

  afterAll(async () => {
    await manager.close();
  });

  test('listFolders returns folders including Inbox', async () => {
    const folders = await client.listFolders();

    expect(Array.isArray(folders)).toBe(true);
    expect(folders.length).toBeGreaterThan(0);

    const inbox = folders.find(f => f.displayName === 'Inbox');
    expect(inbox).toBeDefined();
    expect(typeof inbox!.id).toBe('string');
    expect(typeof inbox!.unreadCount).toBe('number');
    expect(typeof inbox!.totalCount).toBe('number');
    expect(typeof inbox!.childFolderCount).toBe('number');
  }, 40_000);
});
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `npm test -- --testPathPattern=mail.test`
Expected: FAIL — `Cannot find module '../src/mail.js'`

- [ ] **Step 3: Create `src/mail.ts` with `request()`, `listFolders()`, and `normaliseFolder()`**

Create `src/mail.ts`:

```typescript
// src/mail.ts
import type { TokenManager } from './auth.js';
import type {
  MailFolder,
  MailMessage,
  MailAttachment,
  MailAttachmentDownload,
  OwaMailFolder,
  OwaMailMessage,
  OwaMailFolderListResponse,
  OwaMailListResponse,
  OwaMailAttachment,
} from './types.js';

const OWA_BASE = 'https://outlook.office.com/api/v2.0';

export class MailClient {
  constructor(private readonly tokens: TokenManager) {}

  async listFolders(parentFolderId?: string): Promise<MailFolder[]> {
    const path = parentFolderId
      ? `/me/mailfolders/${parentFolderId}/childfolders`
      : '/me/mailfolders';
    const params = new URLSearchParams({
      '$select': 'Id,DisplayName,ParentFolderId,UnreadItemCount,TotalItemCount,ChildFolderCount',
    });
    const res = await this.request('GET', `${path}?${params}`);
    const data = (await res.json()) as OwaMailFolderListResponse;
    return data.value.map(f => this.normaliseFolder(f));
  }

  private async request(
    method: string,
    path: string,
    options: { body?: unknown; headers?: Record<string, string> } = {}
  ): Promise<Response> {
    const token = await this.tokens.getToken();
    const headers: Record<string, string> = {
      Authorization: `Bearer ${token.value}`,
      Accept: 'application/json',
      ...options.headers,
    };
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

  private normaliseFolder(raw: OwaMailFolder): MailFolder {
    return {
      id: raw.Id,
      displayName: raw.DisplayName,
      parentFolderId: raw.ParentFolderId,
      unreadCount: raw.UnreadItemCount,
      totalCount: raw.TotalItemCount,
      childFolderCount: raw.ChildFolderCount,
    };
  }
}
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `npm test -- --testPathPattern=mail.test`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/mail.ts tests/mail.test.ts
git commit -m "feat: MailClient with listFolders method"
```

---

### Task 3: `MailClient.getMessages()`

**Files:**
- Modify: `src/mail.ts`
- Modify: `tests/mail.test.ts`

- [ ] **Step 1: Write the tests**

Add to `tests/mail.test.ts` inside the `describe` block:

```typescript
  test('getMessages returns messages from Inbox', async () => {
    const result = await client.getMessages('Inbox', { limit: 5 });

    expect(Array.isArray(result.messages)).toBe(true);
    expect(result.messages.length).toBeGreaterThan(0);
    expect(result.messages.length).toBeLessThanOrEqual(5);

    const msg = result.messages[0];
    expect(typeof msg.id).toBe('string');
    expect(typeof msg.subject).toBe('string');
    expect(typeof msg.from).toBe('string');
    expect(Array.isArray(msg.toRecipients)).toBe(true);
    expect(typeof msg.receivedDateTime).toBe('string');
    expect(typeof msg.isRead).toBe('boolean');
    expect(typeof msg.hasAttachments).toBe('boolean');
    expect(typeof msg.bodyPreview).toBe('string');
    expect(msg.body).toBeUndefined();
  }, 40_000);

  test('getMessages with unread filter returns only unread', async () => {
    const result = await client.getMessages('Inbox', { filter: 'unread', limit: 5 });

    expect(Array.isArray(result.messages)).toBe(true);
    for (const msg of result.messages) {
      expect(msg.isRead).toBe(false);
    }
  }, 40_000);
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `npm test -- --testPathPattern=mail.test`
Expected: FAIL — `client.getMessages is not a function`

- [ ] **Step 3: Implement `getMessages()` and helpers**

Add to `src/mail.ts` in the `MailClient` class, after `listFolders()`:

```typescript
  async getMessages(
    folderId: string = 'Inbox',
    options: { filter?: string; limit?: number; pageToken?: string } = {}
  ): Promise<{ messages: MailMessage[]; nextPageToken?: string }> {
    const { filter = 'all', limit = 20, pageToken } = options;
    const top = Math.min(limit, 500);

    const params = new URLSearchParams({
      '$select': 'Id,Subject,From,ToRecipients,CcRecipients,ReceivedDateTime,IsRead,HasAttachments,Importance,BodyPreview,ConversationId,Flag,ParentFolderId',
      '$orderby': 'ReceivedDateTime desc',
      '$top': String(top),
    });

    if (pageToken) {
      params.set('$skip', pageToken);
    }

    const filterExpr = this.buildPresetFilter(filter);
    if (filterExpr) {
      params.set('$filter', filterExpr);
    }

    const res = await this.request('GET', `/me/mailfolders/${folderId}/messages?${params}`);
    const data = (await res.json()) as OwaMailListResponse;
    const messages = data.value.map(m => this.normaliseMessage(m));

    const skip = pageToken ? parseInt(pageToken, 10) : 0;
    const nextPageToken = data.value.length >= top ? String(skip + top) : undefined;

    return { messages, nextPageToken };
  }

  private buildPresetFilter(filter: string): string | undefined {
    switch (filter) {
      case 'unread':
        return 'IsRead eq false';
      case 'flagged':
        return "Flag/FlagStatus eq 'Flagged'";
      case 'today': {
        const today = new Date();
        today.setUTCHours(0, 0, 0, 0);
        return `ReceivedDateTime ge ${today.toISOString()}`;
      }
      case 'this_week': {
        const weekAgo = new Date(Date.now() - 7 * 24 * 60 * 60 * 1000);
        weekAgo.setUTCHours(0, 0, 0, 0);
        return `ReceivedDateTime ge ${weekAgo.toISOString()}`;
      }
      case 'all':
      default:
        return undefined;
    }
  }

  private normaliseMessage(raw: OwaMailMessage): MailMessage {
    const msg: MailMessage = {
      id: raw.Id,
      subject: raw.Subject,
      from: this.formatRecipient(raw.From),
      toRecipients: raw.ToRecipients.map(r => this.formatRecipient(r)),
      ccRecipients: raw.CcRecipients.map(r => this.formatRecipient(r)),
      receivedDateTime: raw.ReceivedDateTime,
      isRead: raw.IsRead,
      hasAttachments: raw.HasAttachments,
      importance: raw.Importance,
      bodyPreview: raw.BodyPreview,
      conversationId: raw.ConversationId,
      flag: raw.Flag.FlagStatus,
      parentFolderId: raw.ParentFolderId,
    };
    if (raw.Body) {
      msg.body = raw.Body.Content;
      msg.bodyType = raw.Body.ContentType.toLowerCase() === 'html' ? 'html' : 'text';
    }
    if (raw.Attachments) {
      msg.attachments = raw.Attachments.map(a => ({
        id: a.Id,
        name: a.Name,
        contentType: a.ContentType,
        size: a.Size,
      }));
    }
    return msg;
  }

  private formatRecipient(r: { EmailAddress: { Name: string; Address: string } }): string {
    const { Name, Address } = r.EmailAddress;
    if (!Name || Name === Address) return Address;
    return `${Name} <${Address}>`;
  }
```

- [ ] **Step 4: Run the tests to verify they pass**

Run: `npm test -- --testPathPattern=mail.test`
Expected: PASS (all 3 tests)

- [ ] **Step 5: Commit**

```bash
git add src/mail.ts tests/mail.test.ts
git commit -m "feat: MailClient.getMessages with preset filters and pagination"
```

---

### Task 4: `MailClient.searchMessages()`

**Files:**
- Modify: `src/mail.ts`
- Modify: `tests/mail.test.ts`

- [ ] **Step 1: Write the tests**

Add to `tests/mail.test.ts` inside the `describe` block:

```typescript
  test('searchMessages with query returns results', async () => {
    const result = await client.searchMessages({ query: 'meeting', limit: 5 });

    expect(Array.isArray(result.messages)).toBe(true);
    // Full-text search should find something in a real mailbox
    if (result.messages.length > 0) {
      const msg = result.messages[0];
      expect(typeof msg.id).toBe('string');
      expect(typeof msg.subject).toBe('string');
    }
  }, 40_000);

  test('searchMessages with structured filters', async () => {
    const result = await client.searchMessages({
      receivedAfter: new Date(Date.now() - 30 * 24 * 60 * 60 * 1000).toISOString(),
      limit: 5,
    });

    expect(Array.isArray(result.messages)).toBe(true);
  }, 40_000);

  test('searchMessages throws when query and filters both provided', async () => {
    await expect(
      client.searchMessages({ query: 'hello', from: 'test@example.com' })
    ).rejects.toThrow('Cannot combine');
  }, 40_000);
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `npm test -- --testPathPattern=mail.test`
Expected: FAIL — `client.searchMessages is not a function`

- [ ] **Step 3: Implement `searchMessages()`**

Add to `src/mail.ts` in the `MailClient` class, after `getMessages()`:

```typescript
  async searchMessages(
    options: {
      query?: string;
      from?: string;
      subject?: string;
      hasAttachments?: boolean;
      receivedAfter?: string;
      receivedBefore?: string;
      folderId?: string;
      limit?: number;
    } = {}
  ): Promise<{ messages: MailMessage[]; nextPageToken?: string }> {
    const { query, from, subject, hasAttachments, receivedAfter, receivedBefore, folderId, limit = 20 } = options;
    const top = Math.min(limit, 500);

    const hasStructuredFilters = from !== undefined || subject !== undefined
      || hasAttachments !== undefined || receivedAfter !== undefined || receivedBefore !== undefined;

    if (query && hasStructuredFilters) {
      throw new Error('Cannot combine full-text query with structured filters. Use one or the other.');
    }

    const select = 'Id,Subject,From,ToRecipients,CcRecipients,ReceivedDateTime,IsRead,HasAttachments,Importance,BodyPreview,ConversationId,Flag,ParentFolderId';

    const params = new URLSearchParams({
      '$select': select,
      '$top': String(top),
    });

    const basePath = folderId ? `/me/mailfolders/${folderId}/messages` : '/me/messages';

    if (query) {
      params.set('$search', `"${query}"`);
    } else {
      params.set('$orderby', 'ReceivedDateTime desc');
      const filters: string[] = [];
      if (from) filters.push(`From/EmailAddress/Address eq '${from}'`);
      if (subject) filters.push(`contains(Subject, '${subject}')`);
      if (hasAttachments !== undefined) filters.push(`HasAttachments eq ${hasAttachments}`);
      if (receivedAfter) filters.push(`ReceivedDateTime ge ${receivedAfter}`);
      if (receivedBefore) filters.push(`ReceivedDateTime le ${receivedBefore}`);
      if (filters.length > 0) {
        params.set('$filter', filters.join(' and '));
      }
    }

    const res = await this.request('GET', `${basePath}?${params}`);
    const data = (await res.json()) as OwaMailListResponse;
    const messages = data.value.map(m => this.normaliseMessage(m));

    return { messages };
  }
```

- [ ] **Step 4: Run the tests to verify they pass**

Run: `npm test -- --testPathPattern=mail.test`
Expected: PASS (all 6 tests)

- [ ] **Step 5: Commit**

```bash
git add src/mail.ts tests/mail.test.ts
git commit -m "feat: MailClient.searchMessages with full-text and structured filters"
```

---

### Task 5: `MailClient.getMessage()`

**Files:**
- Modify: `src/mail.ts`
- Modify: `tests/mail.test.ts`

- [ ] **Step 1: Write the tests**

Add to `tests/mail.test.ts` inside the `describe` block:

```typescript
  test('getMessage returns full message with text body', async () => {
    // Get a message ID from Inbox first
    const list = await client.getMessages('Inbox', { limit: 1 });
    expect(list.messages.length).toBeGreaterThan(0);
    const messageId = list.messages[0].id;

    const msg = await client.getMessage(messageId, 'text');

    expect(msg.id).toBe(messageId);
    expect(typeof msg.body).toBe('string');
    expect(msg.bodyType).toBe('text');
    expect(typeof msg.subject).toBe('string');
    expect(typeof msg.from).toBe('string');
  }, 40_000);

  test('getMessage returns HTML body when requested', async () => {
    const list = await client.getMessages('Inbox', { limit: 1 });
    expect(list.messages.length).toBeGreaterThan(0);
    const messageId = list.messages[0].id;

    const msg = await client.getMessage(messageId, 'html');

    expect(msg.id).toBe(messageId);
    expect(typeof msg.body).toBe('string');
    expect(msg.bodyType).toBe('html');
  }, 40_000);
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `npm test -- --testPathPattern=mail.test`
Expected: FAIL — `client.getMessage is not a function`

- [ ] **Step 3: Implement `getMessage()`**

Add to `src/mail.ts` in the `MailClient` class, after `searchMessages()`:

```typescript
  async getMessage(messageId: string, format: 'text' | 'html' = 'text'): Promise<MailMessage> {
    const params = new URLSearchParams({
      '$select': 'Id,Subject,From,ToRecipients,CcRecipients,ReceivedDateTime,IsRead,HasAttachments,Importance,BodyPreview,Body,ConversationId,Flag,ParentFolderId',
      '$expand': 'Attachments($select=Id,Name,ContentType,Size)',
    });
    const res = await this.request('GET', `/me/messages/${messageId}?${params}`, {
      headers: {
        'Prefer': `outlook.body-content-type="${format}"`,
      },
    });
    const raw = (await res.json()) as OwaMailMessage;
    return this.normaliseMessage(raw);
  }
```

- [ ] **Step 4: Run the tests to verify they pass**

Run: `npm test -- --testPathPattern=mail.test`
Expected: PASS (all 8 tests)

- [ ] **Step 5: Commit**

```bash
git add src/mail.ts tests/mail.test.ts
git commit -m "feat: MailClient.getMessage with text/html format support"
```

---

### Task 6: `MailClient.getAttachment()`

**Files:**
- Modify: `src/mail.ts`
- Modify: `tests/mail.test.ts`

- [ ] **Step 1: Write the test**

Add to `tests/mail.test.ts` inside the `describe` block:

```typescript
  test('getAttachment downloads file to disk', async () => {
    // Find a message with attachments
    const result = await client.searchMessages({ hasAttachments: true, limit: 5 });
    const msgWithAttachment = result.messages.find(m => m.hasAttachments);
    if (!msgWithAttachment) {
      console.warn('No messages with attachments found, skipping test');
      return;
    }

    // Get full message to get attachment metadata
    const fullMsg = await client.getMessage(msgWithAttachment.id, 'text');
    expect(fullMsg.attachments).toBeDefined();
    expect(fullMsg.attachments!.length).toBeGreaterThan(0);

    const att = fullMsg.attachments![0];
    const downloaded = await client.getAttachment(fullMsg.id, att.id);

    expect(downloaded.name).toBe(att.name);
    expect(downloaded.contentType).toBe(att.contentType);
    expect(typeof downloaded.filePath).toBe('string');

    // Verify file exists
    const fs = await import('fs');
    expect(fs.existsSync(downloaded.filePath)).toBe(true);

    // Cleanup
    fs.unlinkSync(downloaded.filePath);
  }, 40_000);
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `npm test -- --testPathPattern=mail.test`
Expected: FAIL — `client.getAttachment is not a function`

- [ ] **Step 3: Implement `getAttachment()`**

Add these imports at the top of `src/mail.ts`:

```typescript
import fs from 'fs';
import path from 'path';
import os from 'os';
```

Add to `src/mail.ts` in the `MailClient` class, after `getMessage()`:

```typescript
  async getAttachment(messageId: string, attachmentId: string): Promise<MailAttachmentDownload> {
    const res = await this.request('GET', `/me/messages/${messageId}/attachments/${attachmentId}`);
    const raw = (await res.json()) as OwaMailAttachment;

    const dir = path.join(os.tmpdir(), 'owa-mcp-attachments');
    fs.mkdirSync(dir, { recursive: true });

    const sanitised = raw.Name
      .replace(/[/\\]/g, '_')
      .replace(/[\x00-\x1f]/g, '')
      .slice(0, 200);
    const filePath = path.join(dir, sanitised);

    const buffer = Buffer.from(raw.ContentBytes!, 'base64');
    fs.writeFileSync(filePath, buffer);

    return {
      id: raw.Id,
      name: raw.Name,
      contentType: raw.ContentType,
      size: raw.Size,
      filePath,
    };
  }
```

- [ ] **Step 4: Run the tests to verify they pass**

Run: `npm test -- --testPathPattern=mail.test`
Expected: PASS (all 9 tests)

- [ ] **Step 5: Commit**

```bash
git add src/mail.ts tests/mail.test.ts
git commit -m "feat: MailClient.getAttachment downloads to disk"
```

---

### Task 7: Register all 5 mail tools in `src/index.ts`

**Files:**
- Modify: `src/index.ts`

- [ ] **Step 1: Add import and instantiation**

At the top of `src/index.ts`, add the `MailClient` import alongside the existing `CalendarClient` import:

```typescript
import { MailClient } from './mail.js';
```

After the existing `const calendarClient = new CalendarClient(tokenManager);` line, add:

```typescript
const mailClient = new MailClient(tokenManager);
```

- [ ] **Step 2: Register `list_mail_folders` tool**

Add after the last calendar tool registration (the `follow_calendar_event` tool), before the `async function main()`:

```typescript
server.tool(
  'list_mail_folders',
  'List all mail folders in the mailbox, or child folders of a specific folder.',
  {
    parentFolderId: z.string().optional()
      .describe('List children of this folder. If omitted, lists top-level folders.'),
  },
  async ({ parentFolderId }) => {
    const folders = await mailClient.listFolders(parentFolderId);
    return { content: [{ type: 'text', text: JSON.stringify(folders, null, 2) }] };
  }
);
```

- [ ] **Step 3: Register `get_emails` tool**

```typescript
server.tool(
  'get_emails',
  'Get emails from a specific mailbox folder with optional filtering.',
  {
    folderId: z.string().optional().default('Inbox')
      .describe('Folder ID or well-known name (Inbox, Drafts, SentItems, DeletedItems). Default: Inbox'),
    filter: z.enum(['all', 'unread', 'flagged', 'today', 'this_week']).optional().default('all')
      .describe('Filter preset'),
    limit: z.number().int().min(1).max(500).optional().default(20)
      .describe('Maximum number of emails to return (default 20, max 500)'),
    pageToken: z.string().optional()
      .describe('Pagination token from previous response'),
  },
  async ({ folderId, filter, limit, pageToken }) => {
    const result = await mailClient.getMessages(folderId, { filter, limit, pageToken });
    return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
  }
);
```

- [ ] **Step 4: Register `search_emails` tool**

```typescript
server.tool(
  'search_emails',
  'Search emails using full-text query OR structured filters (mutually exclusive). Use query for natural search, or structured filters for precise field matching.',
  {
    query: z.string().optional()
      .describe('Full-text search query (uses Exchange search index). Cannot be combined with structured filters.'),
    from: z.string().optional()
      .describe('Filter by sender email address'),
    subject: z.string().optional()
      .describe('Filter by subject (contains match)'),
    hasAttachments: z.boolean().optional()
      .describe('Filter by attachment presence'),
    receivedAfter: z.string().optional()
      .describe('ISO 8601 datetime — only messages received after this time'),
    receivedBefore: z.string().optional()
      .describe('ISO 8601 datetime — only messages received before this time'),
    folderId: z.string().optional()
      .describe('Scope search to a specific folder'),
    limit: z.number().int().min(1).max(500).optional().default(20)
      .describe('Maximum results (default 20, max 500)'),
  },
  async (params) => {
    const result = await mailClient.searchMessages(params);
    return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
  }
);
```

- [ ] **Step 5: Register `get_email` tool**

```typescript
server.tool(
  'get_email',
  'Read a single email with full body content and attachment metadata.',
  {
    messageId: z.string().describe('Message ID from get_emails or search_emails'),
    format: z.enum(['text', 'html']).optional().default('text')
      .describe('Body format: "text" (default) or "html"'),
  },
  async ({ messageId, format }) => {
    const message = await mailClient.getMessage(messageId, format);
    return { content: [{ type: 'text', text: JSON.stringify(message, null, 2) }] };
  }
);
```

- [ ] **Step 6: Register `get_attachment` tool**

```typescript
server.tool(
  'get_attachment',
  'Download an email attachment to disk. Returns the file path.',
  {
    messageId: z.string().describe('Message ID'),
    attachmentId: z.string().describe('Attachment ID from get_email response'),
  },
  async ({ messageId, attachmentId }) => {
    const result = await mailClient.getAttachment(messageId, attachmentId);
    return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
  }
);
```

- [ ] **Step 7: Verify it compiles**

Run: `npx tsc --noEmit`
Expected: no errors

- [ ] **Step 8: Run all tests**

Run: `npm test`
Expected: all tests pass (calendar + mail)

- [ ] **Step 9: Commit**

```bash
git add src/index.ts
git commit -m "feat: register 5 mail read tools — folders, emails, search, message, attachment"
```

---

### Task 8: Update CLAUDE.md and package.json description

**Files:**
- Modify: `CLAUDE.md`
- Modify: `package.json`

- [ ] **Step 1: Update `package.json` description**

Change the `description` field in `package.json` from:

```
"description": "MCP server for Microsoft Outlook calendar via Playwright Edge session interception",
```

to:

```
"description": "MCP server for Microsoft Outlook calendar and email via Playwright Edge session interception",
```

Also add `"email"` and `"mail"` to the `keywords` array.

- [ ] **Step 2: Update CLAUDE.md**

In `CLAUDE.md`, update the "What this project is" section to mention email support.

Add `src/mail.ts` to the Key files table:

```
| `src/mail.ts` | `MailClient` — calls `outlook.office.com/api/v2.0` mail endpoints |
```

- [ ] **Step 3: Commit**

```bash
git add CLAUDE.md package.json
git commit -m "docs: update project description to include email support"
```
