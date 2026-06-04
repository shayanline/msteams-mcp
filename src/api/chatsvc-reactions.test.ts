import { describe, it, expect, vi, beforeEach } from 'vitest';
import { ok, err } from '../types/result.js';
import { ErrorCode, createError } from '../types/errors.js';

vi.mock('../utils/http.js', () => ({ httpRequest: vi.fn() }));
vi.mock('../utils/auth-guards.js', () => ({ requireMessageAuthWithConfig: vi.fn() }));

import { httpRequest } from '../utils/http.js';
import { requireMessageAuthWithConfig } from '../utils/auth-guards.js';
import { addReaction, removeReaction } from './chatsvc-reactions.js';

const mockHttp = vi.mocked(httpRequest);
const mockAuth = vi.mocked(requireMessageAuthWithConfig);
const AUTH = ok({ auth: { skypeToken: 'sk', authToken: 'at', userMri: '8:orgid:me' }, region: 'uk', baseUrl: 'https://teams.microsoft.com' });
const httpOk = (data: unknown) => ok({ status: 200, headers: new Headers(), data } as never);
const lastCall = () => mockHttp.mock.calls.at(-1)!;
const lastBody = () => JSON.parse((lastCall()[1] as { body: string }).body);

beforeEach(() => {
  vi.clearAllMocks();
  mockAuth.mockReturnValue(AUTH as never);
  mockHttp.mockResolvedValue(httpOk(null));
});

describe('addReaction', () => {
  it('issues a PUT with the emoji key and a value timestamp', async () => {
    const now = 1_700_000_000_000;
    vi.spyOn(Date, 'now').mockReturnValue(now);
    const res = await addReaction('19:c@thread.v2', 'm1', 'like');
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value).toEqual({ conversationId: '19:c@thread.v2', messageId: 'm1', emoji: 'like' });
    expect(lastCall()[1]).toMatchObject({ method: 'PUT' });
    expect(lastBody()).toEqual({ emotions: { key: 'like', value: now } });
  });

  it('propagates an auth failure without calling http', async () => {
    mockAuth.mockReturnValue(err(createError(ErrorCode.AUTH_REQUIRED, 'no auth')) as never);
    const res = await addReaction('19:c', 'm1', 'like');
    expect(res.ok).toBe(false);
    expect(mockHttp).not.toHaveBeenCalled();
  });

  it('propagates an http failure', async () => {
    mockHttp.mockResolvedValueOnce(err(createError(ErrorCode.NETWORK_ERROR, 'boom')) as never);
    const res = await addReaction('19:c', 'm1', 'like');
    expect(res.ok).toBe(false);
  });
});

describe('removeReaction', () => {
  it('issues a DELETE with the emoji key and no value', async () => {
    const res = await removeReaction('19:c@thread.v2', 'm2', 'heart');
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value).toEqual({ conversationId: '19:c@thread.v2', messageId: 'm2', emoji: 'heart' });
    expect(lastCall()[1]).toMatchObject({ method: 'DELETE' });
    expect(lastBody()).toEqual({ emotions: { key: 'heart' } });
  });
});
