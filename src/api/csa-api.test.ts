import { describe, it, expect, vi, beforeEach } from 'vitest';
import { ok } from '../types/result.js';

vi.mock('../utils/http.js', () => ({ httpRequest: vi.fn() }));
vi.mock('../utils/auth-guards.js', () => ({ requireCsaAuth: vi.fn(), getApiConfig: vi.fn() }));
vi.mock('./chatsvc-api.js', () => ({ getConversationProperties: vi.fn(), extractParticipantNames: vi.fn() }));

import { httpRequest } from '../utils/http.js';
import { requireCsaAuth, getApiConfig } from '../utils/auth-guards.js';
import { getConversationProperties, extractParticipantNames } from './chatsvc-api.js';
import { getFavorites, addFavorite, getCustomEmojis, getMyTeamsAndChannels } from './csa-api.js';

const mockHttp = vi.mocked(httpRequest);
const mockAuth = vi.mocked(requireCsaAuth);
const mockConfig = vi.mocked(getApiConfig);
const mockProps = vi.mocked(getConversationProperties);
const mockNames = vi.mocked(extractParticipantNames);
const httpOk = (data: unknown) => ok({ status: 200, headers: new Headers(), data } as never);

beforeEach(() => {
  vi.clearAllMocks();
  mockAuth.mockReturnValue(ok({ auth: { skypeToken: 'sk', authToken: 'at', userMri: '8:orgid:me' }, csaToken: 'csa' }) as never);
  mockConfig.mockReturnValue({ region: 'uk', baseUrl: 'https://teams.microsoft.com' } as never);
  mockProps.mockResolvedValue(ok({ displayName: 'Chat X', conversationType: 'chat' }) as never);
  mockNames.mockResolvedValue(ok('Alice, Bob') as never);
});

describe('getFavorites', () => {
  it('extracts the Favorites folder items and enriches names', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({
      folderHierarchyVersion: 5,
      conversationFolders: [
        { folderType: 'Favorites', id: 'fid', conversationFolderItems: [{ conversationId: '19:x', createdTime: 1 }] },
      ],
    }));
    const res = await getFavorites();
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.folderId).toBe('fid');
    expect(res.value.favorites[0]).toMatchObject({ conversationId: '19:x', displayName: 'Chat X', conversationType: 'chat' });
  });

  it('returns empty when there is no Favorites folder', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ folderHierarchyVersion: 2, conversationFolders: [] }));
    const res = await getFavorites();
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.favorites).toEqual([]);
  });

  it('falls back to participant names when no display name', async () => {
    mockProps.mockResolvedValue(ok({ displayName: undefined, conversationType: 'chat' }) as never);
    mockHttp.mockResolvedValueOnce(httpOk({ conversationFolders: [{ folderType: 'Favorites', id: 'fid', conversationFolderItems: [{ conversationId: '19:y' }] }] }));
    const res = await getFavorites();
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.favorites[0].displayName).toBe('Alice, Bob');
  });
});

describe('addFavorite', () => {
  it('reads current state then POSTs an AddItem action', async () => {
    // first call: getFavorites GET; second call: modify POST
    mockHttp
      .mockResolvedValueOnce(httpOk({ folderHierarchyVersion: 7, conversationFolders: [{ folderType: 'Favorites', id: 'fid', conversationFolderItems: [] }] }))
      .mockResolvedValueOnce(httpOk(null));
    const res = await addFavorite('19:z');
    expect(res.ok).toBe(true);
    const post = mockHttp.mock.calls[1];
    expect(post[1]).toMatchObject({ method: 'POST' });
    const sent = JSON.parse((post[1] as { body: string }).body);
    expect(sent.folderHierarchyVersion).toBe(7);
    expect(sent.actions[0]).toMatchObject({ action: 'AddItem', folderId: 'fid', itemId: '19:z' });
  });
});

describe('getCustomEmojis', () => {
  it('flattens categories and skips deleted emoticons', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ categories: [
      { emoticons: [
        { id: 'e1;abc', shortcuts: ['smile'], description: 'Smile' },
        { id: 'e2;def', isDeleted: true },
        { id: 'e3;ghi' },
      ] },
    ] }));
    const res = await getCustomEmojis();
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.emojis).toEqual([
      { id: 'e1;abc', shortcut: 'smile', description: 'Smile', createdOn: undefined },
      { id: 'e3;ghi', shortcut: 'e3', description: 'e3', createdOn: undefined },
    ]);
  });
});

describe('getMyTeamsAndChannels', () => {
  it('returns parsed teams (empty for empty payload)', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ teams: [] }));
    const res = await getMyTeamsAndChannels();
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(Array.isArray(res.value.teams)).toBe(true);
  });

  it('propagates auth failure', async () => {
    mockAuth.mockReturnValueOnce({ ok: false, error: { code: 'AUTH_REQUIRED' } } as never);
    const res = await getMyTeamsAndChannels();
    expect(res.ok).toBe(false);
    expect(mockHttp).not.toHaveBeenCalled();
  });
});
