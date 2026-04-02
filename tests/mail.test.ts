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
});
