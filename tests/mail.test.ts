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
