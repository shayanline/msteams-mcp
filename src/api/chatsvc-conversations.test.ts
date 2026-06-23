import { describe, it, expect, vi, beforeEach } from 'vitest';
import { ok } from '../types/result.js';

vi.mock('../utils/http.js', () => ({ httpRequest: vi.fn() }));
vi.mock('../utils/auth-guards.js', () => ({ requireMessageAuthWithConfig: vi.fn() }));
vi.mock('./chatsvc-messaging.js', () => ({
  getMessage: vi.fn(),
  sendMessage: vi.fn(),
  buildReplyQuoteHtml: vi.fn(() => '<quote/>'),
}));

import { httpRequest } from '../utils/http.js';
import { requireMessageAuthWithConfig } from '../utils/auth-guards.js';
import { getMessage, sendMessage } from './chatsvc-messaging.js';
import {
  getChatMembers, addMember, removeMember, leaveChat, renameChat,
  pinMessage, unpinMessage, setMuted, forwardMessage,
} from './chatsvc-conversations.js';

const mockHttp = vi.mocked(httpRequest);
const mockAuth = vi.mocked(requireMessageAuthWithConfig);
const mockGetMessage = vi.mocked(getMessage);
const mockSend = vi.mocked(sendMessage);

const AUTH = ok({ auth: { skypeToken: 'sk', authToken: 'at', userMri: '8:orgid:me' }, region: 'uk', baseUrl: 'https://teams.microsoft.com' });
const httpOk = (data: unknown = null) => ok({ status: 200, headers: new Headers(), data } as never);
const lastUrl = () => mockHttp.mock.calls.at(-1)![0] as string;
const lastInit = () => mockHttp.mock.calls.at(-1)![1] as { method: string; body?: string };

beforeEach(() => {
  vi.clearAllMocks();
  mockAuth.mockReturnValue(AUTH as never);
});

describe('getChatMembers', () => {
  it('parses members and flags the current user', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ totalMemberCount: 2, members: [
      { id: '8:orgid:me', role: 'Admin' },
      { id: '8:orgid:other', role: 'User' },
    ] }));
    const res = await getChatMembers('19:abc@thread.v2');
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.totalMemberCount).toBe(2);
    expect(res.value.members).toEqual([
      { mri: '8:orgid:me', role: 'Admin', isCurrentUser: true },
      { mri: '8:orgid:other', role: 'User', isCurrentUser: false },
    ]);
    expect(lastUrl()).toContain('/threads/19%3Aabc%40thread.v2/members');
    expect(lastInit().method).toBe('GET');
  });
});

describe('addMember', () => {
  it('adds an MRI member with a role via PUT', async () => {
    mockHttp.mockResolvedValueOnce(httpOk());
    const res = await addMember('19:c@thread.v2', '8:orgid:peer', 'User');
    expect(res.ok).toBe(true);
    expect(lastInit().method).toBe('PUT');
    expect(lastUrl()).toContain('/members/8%3Aorgid%3Apeer');
    expect(JSON.parse(lastInit().body!)).toEqual({ role: 'User' });
  });

  it('normalises a raw GUID to an MRI', async () => {
    mockHttp.mockResolvedValueOnce(httpOk());
    await addMember('19:c@thread.v2', 'dde37b63-a4ac-4edb-a7b0-385263022300');
    expect(lastUrl()).toContain('8%3Aorgid%3Adde37b63-a4ac-4edb-a7b0-385263022300');
  });

  it('rejects an invalid identifier without calling http', async () => {
    const res = await addMember('19:c@thread.v2', 'not-an-id');
    expect(res.ok).toBe(false);
    expect(mockHttp).not.toHaveBeenCalled();
  });
});

describe('removeMember / leaveChat', () => {
  it('removes a member via DELETE', async () => {
    mockHttp.mockResolvedValueOnce(httpOk());
    const res = await removeMember('19:c@thread.v2', '8:orgid:peer');
    expect(res.ok).toBe(true);
    expect(lastInit().method).toBe('DELETE');
    expect(lastUrl()).toContain('/members/8%3Aorgid%3Apeer');
  });

  it('leaveChat deletes the current user', async () => {
    mockHttp.mockResolvedValueOnce(httpOk());
    const res = await leaveChat('19:c@thread.v2');
    expect(res.ok).toBe(true);
    expect(lastInit().method).toBe('DELETE');
    expect(lastUrl()).toContain('/members/8%3Aorgid%3Ame');
  });
});

describe('renameChat', () => {
  it('PUTs the topic thread property', async () => {
    mockHttp.mockResolvedValueOnce(httpOk());
    const res = await renameChat('19:c@thread.v2', 'New name');
    expect(res.ok).toBe(true);
    expect(lastUrl()).toContain('/properties?name=topic');
    expect(JSON.parse(lastInit().body!)).toEqual({ topic: 'New name' });
  });
});

describe('pin / unpin', () => {
  it('pins a message as a JSON-encoded pinnedItems array', async () => {
    mockHttp.mockResolvedValueOnce(httpOk());
    const res = await pinMessage('19:c@thread.v2', '17800001');
    expect(res.ok).toBe(true);
    expect(lastUrl()).toContain('name=pinnedItems');
    const body = JSON.parse(lastInit().body!);
    expect(JSON.parse(body.pinnedItems)).toEqual([{ itemId: '17800001', itemType: 'Message' }]);
  });

  it('unpins by setting an empty array', async () => {
    mockHttp.mockResolvedValueOnce(httpOk());
    await unpinMessage('19:c@thread.v2');
    expect(JSON.parse(JSON.parse(lastInit().body!).pinnedItems)).toEqual([]);
  });
});

describe('setMuted', () => {
  it('mutes by setting alerts=false', async () => {
    mockHttp.mockResolvedValueOnce(httpOk());
    const res = await setMuted('19:c@thread.v2', true);
    expect(res.ok).toBe(true);
    expect(lastUrl()).toContain('name=alerts');
    expect(JSON.parse(lastInit().body!)).toEqual({ alerts: 'false' });
  });

  it('unmutes by setting alerts=true', async () => {
    mockHttp.mockResolvedValueOnce(httpOk());
    await setMuted('19:c@thread.v2', false);
    expect(JSON.parse(lastInit().body!)).toEqual({ alerts: 'true' });
  });
});

describe('forwardMessage', () => {
  it('quotes the original and sends it to the target with a comment', async () => {
    mockGetMessage.mockResolvedValueOnce(ok({ id: '1', content: 'hi', contentType: 'text', sender: { mri: '8:orgid:x' }, timestamp: 't', conversationId: 's' }) as never);
    mockSend.mockResolvedValueOnce(ok({ messageId: 'm', timestamp: 1 }) as never);
    const res = await forwardMessage('19:src@thread.v2', '1', '48:notes', 'FYI');
    expect(res.ok).toBe(true);
    const [target, content, opts] = mockSend.mock.calls[0] as [string, string, { rawContentHtml: string }];
    expect(target).toBe('48:notes');
    expect(content).toBe('');
    // comment is prepended; no hardcoded "Forwarded message:" label
    expect(opts.rawContentHtml).toContain('FYI');
    expect(opts.rawContentHtml).not.toContain('Forwarded message');
    expect(opts.rawContentHtml).toContain('<quote/>');
  });

  it('returns the error when the original cannot be fetched', async () => {
    mockGetMessage.mockResolvedValueOnce({ ok: false, error: { code: 'NOT_FOUND' } } as never);
    const res = await forwardMessage('19:src@thread.v2', '1', '48:notes');
    expect(res.ok).toBe(false);
    expect(mockSend).not.toHaveBeenCalled();
  });

  it('forwards without a comment (no note prefix) and returns the send error', async () => {
    mockGetMessage.mockResolvedValueOnce(ok({ id: '1', content: 'hi', contentType: 'text', sender: { mri: '8:orgid:x' }, timestamp: 't', conversationId: 's' }) as never);
    mockSend.mockResolvedValueOnce({ ok: false, error: { code: 'API_ERROR' } } as never);
    const res = await forwardMessage('19:src@thread.v2', '1', '48:notes');
    expect(res.ok).toBe(false);
    const [, , opts] = mockSend.mock.calls[0] as [string, string, { rawContentHtml: string }];
    expect(opts.rawContentHtml).not.toContain('<br><br>'); // no comment note prefix
  });
});

describe('auth failures', () => {
  const AUTH_ERR = { ok: false, error: { code: 'AUTH_REQUIRED' } } as never;

  it('getChatMembers returns the auth error without calling http', async () => {
    mockAuth.mockReturnValueOnce(AUTH_ERR);
    const res = await getChatMembers('19:c@thread.v2');
    expect(res.ok).toBe(false);
    expect(mockHttp).not.toHaveBeenCalled();
  });

  it('addMember returns the auth error', async () => {
    mockAuth.mockReturnValueOnce(AUTH_ERR);
    expect((await addMember('19:c@thread.v2', '8:orgid:peer')).ok).toBe(false);
    expect(mockHttp).not.toHaveBeenCalled();
  });

  it('removeMember returns the auth error', async () => {
    mockAuth.mockReturnValueOnce(AUTH_ERR);
    expect((await removeMember('19:c@thread.v2', '8:orgid:peer')).ok).toBe(false);
  });

  it('leaveChat returns the auth error', async () => {
    mockAuth.mockReturnValueOnce(AUTH_ERR);
    expect((await leaveChat('19:c@thread.v2')).ok).toBe(false);
  });

  it('renameChat returns the auth error', async () => {
    mockAuth.mockReturnValueOnce(AUTH_ERR);
    expect((await renameChat('19:c@thread.v2', 'x')).ok).toBe(false);
  });

  it('setMuted returns the auth error', async () => {
    mockAuth.mockReturnValueOnce(AUTH_ERR);
    expect((await setMuted('19:c@thread.v2', true)).ok).toBe(false);
  });

  it('pinMessage returns the auth error (via setPinnedItems)', async () => {
    mockAuth.mockReturnValueOnce(AUTH_ERR);
    expect((await pinMessage('19:c@thread.v2', '1')).ok).toBe(false);
    expect(mockHttp).not.toHaveBeenCalled();
  });
});

describe('http failures propagate', () => {
  const httpErr = { ok: false, error: { code: 'API_ERROR' } } as never;

  it('getChatMembers propagates the http error', async () => {
    mockHttp.mockResolvedValueOnce(httpErr);
    expect((await getChatMembers('19:c@thread.v2')).ok).toBe(false);
  });

  it('addMember propagates the http error', async () => {
    mockHttp.mockResolvedValueOnce(httpErr);
    expect((await addMember('19:c@thread.v2', '8:orgid:peer')).ok).toBe(false);
  });

  it('removeMember propagates the http error', async () => {
    mockHttp.mockResolvedValueOnce(httpErr);
    expect((await removeMember('19:c@thread.v2', '8:orgid:peer')).ok).toBe(false);
  });

  it('removeMember rejects an invalid identifier without calling http', async () => {
    const res = await removeMember('19:c@thread.v2', 'not-an-id');
    expect(res.ok).toBe(false);
    expect(mockHttp).not.toHaveBeenCalled();
  });

  it('leaveChat propagates the http error', async () => {
    mockHttp.mockResolvedValueOnce(httpErr);
    expect((await leaveChat('19:c@thread.v2')).ok).toBe(false);
  });

  it('renameChat propagates the http error', async () => {
    mockHttp.mockResolvedValueOnce(httpErr);
    expect((await renameChat('19:c@thread.v2', 'x')).ok).toBe(false);
  });

  it('pinMessage propagates the http error', async () => {
    mockHttp.mockResolvedValueOnce(httpErr);
    expect((await pinMessage('19:c@thread.v2', '1')).ok).toBe(false);
  });

  it('unpinMessage propagates the http error', async () => {
    mockHttp.mockResolvedValueOnce(httpErr);
    expect((await unpinMessage('19:c@thread.v2')).ok).toBe(false);
  });

  it('setMuted propagates the http error', async () => {
    mockHttp.mockResolvedValueOnce(httpErr);
    expect((await setMuted('19:c@thread.v2', false)).ok).toBe(false);
  });
});

describe('getChatMembers defaults', () => {
  it('defaults members to [], role to User and total to member count when absent', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ members: [{ id: '8:orgid:nobody' }] }));
    const res = await getChatMembers('19:c@thread.v2');
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.members).toEqual([{ mri: '8:orgid:nobody', role: 'User', isCurrentUser: false }]);
    expect(res.value.totalMemberCount).toBe(1);
  });

  it('returns an empty member list when the data has no members', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({}));
    const res = await getChatMembers('19:c@thread.v2');
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.members).toEqual([]);
    expect(res.value.totalMemberCount).toBe(0);
  });
});
