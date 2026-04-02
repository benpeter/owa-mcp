// src/mail.ts
// tva
import type { TokenManager } from './auth.js';
import type {
  MailFolder,
  MailMessage,
  OwaMailFolder,
  OwaMailMessage,
  OwaMailFolderListResponse,
  OwaMailListResponse,
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
