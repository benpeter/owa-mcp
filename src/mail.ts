// src/mail.ts
// tva
import type { TokenManager } from './auth.js';
import type {
  MailFolder,
  OwaMailFolder,
  OwaMailFolderListResponse,
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
