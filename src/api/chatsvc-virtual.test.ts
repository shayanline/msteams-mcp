import { describe, it, expect, vi, beforeEach } from 'vitest';
import { ok, err } from '../types/result.js';
import { ErrorCode, createError } from '../types/errors.js';

vi.mock('../utils/http.js', () => ({ httpRequest: vi.fn() }));
vi.mock('../utils/auth-guards.js', () => ({
  requireMessageAuthWithConfig: vi.fn(),
  getTeamsBaseUrl: vi.fn(),
  getTenantId: vi.fn(),
}));
vi.mock('../utils/parsers.js', () => ({
  parseVirtualConversationMessage: vi.fn(),
}));

import { httpRequest } from '../utils/http.js';
import { requireMessageAuthWithConfig, getTeamsBaseUrl, getTenantId } from '../utils/auth-guards.js';
import { parseVirtualConversationMessage } from '../utils/parsers.js';
import { getSavedMessages, getFollowedThreads, saveMessage, unsaveMessage } from './chatsvc-virtual.js';

const mockHttp = vi.mocked(httpRequest);
const mockAuth = vi.mocked(requireMessageAuthWithConfig);
const mockParse = vi.mocked(parseVirtualConversationMessage);
const mockTenant = vi.mocked(getTenantId);
const mockBaseUrl = vi.mocked(getTeamsBaseUrl);
const AUTH = ok({ auth: { skypeToken: 'sk', authToken: 'at', userMri: '8:orgid:me' }, region: 'uk', baseUrl: 'https://teams.microsoft.com' });
const httpOk = (data: unknown) => ok({ status: 200, headers: new Headers(), data } as never);
const lastCall = () => mockHttp.mock.calls.at(-1)!;
const lastUrl = () => lastCall()[0] as string;
const lastBody = () => JSON.parse((lastCall()[1] as { body: string }).body);

const item = (over: Record<string, unknown> = {}) => ({
  id: 'i1',
  content: 'c',
  contentType: 'RichText/Html',
  sender: { mri: '8:orgid:x', displayName: 'X' },
  timestamp: '2024-01-01T00:00:00Z',
  sourceConversationId: '19:src',
  sourceReferenceId: 'ref1',
  messageLink: 'link',
  ...over,
});

beforeEach(() => {
  vi.clearAllMocks();
  mockAuth.mockReturnValue(AUTH as never);
  mockTenant.mockReturnValue('tid');
  mockBaseUrl.mockReturnValue('https://teams.microsoft.com');
});

describe('getSavedMessages', () => {
  it('propagates an auth failure without calling http', async () => {
    mockAuth.mockReturnValue(err(createError(ErrorCode.AUTH_REQUIRED, 'no auth')) as never);
    const res = await getSavedMessages();
    expect(res.ok).toBe(false);
    expect(mockHttp).not.toHaveBeenCalled();
  });

  it('propagates an http failure', async () => {
    mockHttp.mockResolvedValueOnce(err(createError(ErrorCode.NETWORK_ERROR, 'boom')) as never);
    const res = await getSavedMessages();
    expect(res.ok).toBe(false);
  });

  it('returns empty messages when messages is not an array', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ messages: 'nope' }));
    const res = await getSavedMessages();
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.messages).toEqual([]);
  });

  it('parses, skips nulls, maps sourceMessageId and sorts (newest first)', async () => {
    mockParse
      .mockReturnValueOnce(item({ id: 'old', timestamp: '2024-01-01T00:00:00Z', sourceReferenceId: 'r-old' }) as never)
      .mockReturnValueOnce(null as never)
      .mockReturnValueOnce(item({ id: 'new', timestamp: '2024-02-01T00:00:00Z', sourceReferenceId: 'r-new' }) as never);
    mockHttp.mockResolvedValueOnce(httpOk({ messages: [{}, {}, {}] }));
    const res = await getSavedMessages({ limit: 5 });
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(lastUrl()).toContain('pageSize=5');
    expect(res.value.messages.map(m => m.id)).toEqual(['new', 'old']);
    expect(res.value.messages[0].sourceMessageId).toBe('r-new');
  });

  it('uses default page size of 50', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ messages: [] }));
    await getSavedMessages();
    expect(lastUrl()).toContain('pageSize=50');
  });
});

describe('getFollowedThreads', () => {
  it('maps sourcePostId from the reference id', async () => {
    mockParse.mockReturnValueOnce(item({ id: 't1', sourceReferenceId: 'post1' }) as never);
    mockHttp.mockResolvedValueOnce(httpOk({ messages: [{}] }));
    const res = await getFollowedThreads();
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.threads[0].sourcePostId).toBe('post1');
  });

  it('propagates an auth failure', async () => {
    mockAuth.mockReturnValue(err(createError(ErrorCode.AUTH_REQUIRED, 'no auth')) as never);
    const res = await getFollowedThreads();
    expect(res.ok).toBe(false);
  });
});

describe('saveMessage / unsaveMessage', () => {
  it('saveMessage PUTs s:1 with the numeric mid and returns saved true', async () => {
    mockHttp.mockResolvedValueOnce(httpOk(null));
    const res = await saveMessage('19:c@thread.v2', '12345');
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value).toEqual({ conversationId: '19:c@thread.v2', messageId: '12345', saved: true });
    expect(lastCall()[1]).toMatchObject({ method: 'PUT' });
    expect(lastBody()).toEqual({ s: 1, mid: 12345 });
  });

  it('unsaveMessage PUTs s:0 and returns saved false', async () => {
    mockHttp.mockResolvedValueOnce(httpOk(null));
    const res = await unsaveMessage('19:c@thread.v2', '999');
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.saved).toBe(false);
    expect(lastBody()).toEqual({ s: 0, mid: 999 });
  });

  it('uses rootMessageId in the URL path when provided', async () => {
    mockHttp.mockResolvedValueOnce(httpOk(null));
    await saveMessage('19:c@thread.v2', '111', '222');
    expect(lastUrl()).toContain('222');
    expect(lastBody()).toEqual({ s: 1, mid: 111 });
  });

  it('rejects a non-numeric message id without calling http', async () => {
    const res = await saveMessage('19:c', 'abc');
    expect(res.ok).toBe(false);
    if (res.ok) return;
    expect(res.error.code).toBe(ErrorCode.INVALID_INPUT);
    expect(mockHttp).not.toHaveBeenCalled();
  });

  it('propagates an auth failure', async () => {
    mockAuth.mockReturnValue(err(createError(ErrorCode.AUTH_REQUIRED, 'no auth')) as never);
    const res = await saveMessage('19:c', '12');
    expect(res.ok).toBe(false);
    expect(mockHttp).not.toHaveBeenCalled();
  });

  it('propagates an http failure', async () => {
    mockHttp.mockResolvedValueOnce(err(createError(ErrorCode.NETWORK_ERROR, 'boom')) as never);
    const res = await saveMessage('19:c', '12');
    expect(res.ok).toBe(false);
  });
});
