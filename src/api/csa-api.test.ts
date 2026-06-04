import { describe, it, expect, vi, beforeEach } from 'vitest';
import { ok } from '../types/result.js';

vi.mock('../utils/http.js', () => ({ httpRequest: vi.fn() }));
vi.mock('../utils/auth-guards.js', () => ({ requireCsaAuth: vi.fn(), getApiConfig: vi.fn() }));
vi.mock('./chatsvc-api.js', () => ({ getConversationProperties: vi.fn(), extractParticipantNames: vi.fn() }));

import { httpRequest } from '../utils/http.js';
import { requireCsaAuth, getApiConfig } from '../utils/auth-guards.js';
import { getConversationProperties, extractParticipantNames } from './chatsvc-api.js';
import { getFavorites, addFavorite, removeFavorite, getCustomEmojis, getMyTeamsAndChannels } from './csa-api.js';

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

  it('propagates an http failure', async () => {
    mockHttp.mockResolvedValueOnce({ ok: false, error: { code: 'API_ERROR' } } as never);
    const res = await getMyTeamsAndChannels();
    expect(res.ok).toBe(false);
  });
});

describe('getFavorites error paths and enrichment fallbacks', () => {
  it('propagates auth failure without calling http', async () => {
    mockAuth.mockReturnValueOnce({ ok: false, error: { code: 'AUTH_REQUIRED' } } as never);
    const res = await getFavorites();
    expect(res.ok).toBe(false);
    expect(mockHttp).not.toHaveBeenCalled();
  });

  it('propagates an http failure', async () => {
    mockHttp.mockResolvedValueOnce({ ok: false, error: { code: 'API_ERROR' } } as never);
    const res = await getFavorites();
    expect(res.ok).toBe(false);
  });

  it('handles a Favorites folder with no items', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ conversationFolders: [{ folderType: 'Favorites', id: 'fid' }] }));
    const res = await getFavorites();
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.favorites).toEqual([]);
    expect(res.value.folderId).toBe('fid');
  });

  it('leaves the display name unset when both enrichment sources fail', async () => {
    mockProps.mockResolvedValue({ ok: false, error: { code: 'NOT_FOUND' } } as never);
    mockNames.mockResolvedValue({ ok: false, error: { code: 'NOT_FOUND' } } as never);
    mockHttp.mockResolvedValueOnce(httpOk({ conversationFolders: [{ folderType: 'Favorites', id: 'fid', conversationFolderItems: [{ conversationId: '19:z' }] }] }));
    const res = await getFavorites();
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.favorites[0].displayName).toBeUndefined();
  });
});

describe('modifyFavorite error paths', () => {
  it('removeFavorite reads state then POSTs a RemoveItem action', async () => {
    mockHttp
      .mockResolvedValueOnce(httpOk({ folderHierarchyVersion: 3, conversationFolders: [{ folderType: 'Favorites', id: 'fid', conversationFolderItems: [] }] }))
      .mockResolvedValueOnce(httpOk(null));
    const res = await removeFavorite('19:z');
    expect(res.ok).toBe(true);
    const sent = JSON.parse((mockHttp.mock.calls[1][1] as { body: string }).body);
    expect(sent.actions[0]).toMatchObject({ action: 'RemoveItem', folderId: 'fid', itemId: '19:z' });
  });

  it('returns the auth error before reading state', async () => {
    mockAuth.mockReturnValueOnce({ ok: false, error: { code: 'AUTH_REQUIRED' } } as never);
    const res = await addFavorite('19:z');
    expect(res.ok).toBe(false);
    expect(mockHttp).not.toHaveBeenCalled();
  });

  it('returns the error when reading current state fails', async () => {
    mockHttp.mockResolvedValueOnce({ ok: false, error: { code: 'API_ERROR' } } as never);
    const res = await addFavorite('19:z');
    expect(res.ok).toBe(false);
  });

  it('errors when there is no Favorites folder to modify', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ conversationFolders: [] }));
    const res = await addFavorite('19:z');
    expect(res.ok).toBe(false);
  });

  it('propagates a failure from the POST modification', async () => {
    mockHttp
      .mockResolvedValueOnce(httpOk({ folderHierarchyVersion: 1, conversationFolders: [{ folderType: 'Favorites', id: 'fid', conversationFolderItems: [] }] }))
      .mockResolvedValueOnce({ ok: false, error: { code: 'API_ERROR' } } as never);
    const res = await addFavorite('19:z');
    expect(res.ok).toBe(false);
  });
});

describe('getCustomEmojis error paths and empty shapes', () => {
  it('propagates auth failure', async () => {
    mockAuth.mockReturnValueOnce({ ok: false, error: { code: 'AUTH_REQUIRED' } } as never);
    const res = await getCustomEmojis();
    expect(res.ok).toBe(false);
    expect(mockHttp).not.toHaveBeenCalled();
  });

  it('propagates an http failure', async () => {
    mockHttp.mockResolvedValueOnce({ ok: false, error: { code: 'API_ERROR' } } as never);
    const res = await getCustomEmojis();
    expect(res.ok).toBe(false);
  });

  it('returns an empty list when there are no categories', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({}));
    const res = await getCustomEmojis();
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.emojis).toEqual([]);
  });

  it('skips a category that has no emoticons', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ categories: [{}, { emoticons: [{ id: 'e1;x', shortcuts: ['s'], description: 'd' }] }] }));
    const res = await getCustomEmojis();
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.emojis).toEqual([{ id: 'e1;x', shortcut: 's', description: 'd', createdOn: undefined }]);
  });
});
