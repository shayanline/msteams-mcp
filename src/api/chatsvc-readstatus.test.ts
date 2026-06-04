import { describe, it, expect, vi, beforeEach } from 'vitest';
import { ok } from '../types/result.js';

vi.mock('../utils/http.js', () => ({ httpRequest: vi.fn() }));
vi.mock('../utils/auth-guards.js', () => ({ requireMessageAuthWithConfig: vi.fn() }));
vi.mock('./chatsvc-messaging.js', () => ({ getThreadMessages: vi.fn() }));

import { httpRequest } from '../utils/http.js';
import { requireMessageAuthWithConfig } from '../utils/auth-guards.js';
import { getThreadMessages } from './chatsvc-messaging.js';
import {
  getConsumptionHorizon, markAsRead, markUnread, getUnreadStatus,
  getUnreadConversations, listConversations,
} from './chatsvc-readstatus.js';

const mockHttp = vi.mocked(httpRequest);
const mockAuth = vi.mocked(requireMessageAuthWithConfig);
const mockThread = vi.mocked(getThreadMessages);
const AUTH = ok({ auth: { skypeToken: 'sk', authToken: 'at', userMri: '8:orgid:me' }, region: 'uk', baseUrl: 'https://teams.microsoft.com' });
const httpOk = (data: unknown) => ok({ status: 200, headers: new Headers(), data } as never);
const body = () => JSON.parse((mockHttp.mock.calls.at(-1)![1] as { body: string }).body);

beforeEach(() => {
  vi.clearAllMocks();
  mockAuth.mockReturnValue(AUTH as never);
});

describe('getConsumptionHorizon', () => {
  it('finds the current user horizon and parses it', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ version: 'v1', consumptionhorizons: [
      { id: '8:orgid:other', consumptionhorizon: '1;2;mOther' },
      { id: '8:orgid:me', consumptionhorizon: '100;200;mX' },
    ] }));
    const res = await getConsumptionHorizon('19:c@thread.v2');
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.lastReadMessageId).toBe('mX');
    expect(res.value.lastReadTimestamp).toBe(100);
    expect(res.value.consumptionHorizons).toHaveLength(2);
  });
});

describe('markAsRead / markUnread', () => {
  it('markAsRead sets an equal triple horizon', async () => {
    mockHttp.mockResolvedValueOnce(httpOk(null));
    const res = await markAsRead('19:c@thread.v2', '1780000000000');
    expect(res.ok).toBe(true);
    expect(body()).toEqual({ consumptionhorizon: '1780000000000;1780000000000;1780000000000' });
  });

  it('markUnread moves the horizon one ms before the message', async () => {
    mockHttp.mockResolvedValueOnce(httpOk(null));
    const res = await markUnread('19:c@thread.v2', '1780000000000');
    expect(res.ok).toBe(true);
    expect(body()).toEqual({ consumptionhorizon: '1779999999999;1779999999999;1779999999999' });
  });

  it('markUnread falls back to the id when it is not numeric', async () => {
    mockHttp.mockResolvedValueOnce(httpOk(null));
    await markUnread('19:c@thread.v2', 'abc');
    expect(body()).toEqual({ consumptionhorizon: 'abc;abc;abc' });
  });
});

describe('getUnreadStatus', () => {
  it('counts messages from others after the last read position', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ consumptionhorizons: [{ id: '8:orgid:me', consumptionhorizon: '1;2;m1' }] }));
    mockThread.mockResolvedValueOnce(ok({ conversationId: '19:c@thread.v2', messages: [
      { id: 'm3', isFromMe: false },
      { id: 'm2', isFromMe: false },
      { id: 'm1', isFromMe: false },
    ] }) as never);
    const res = await getUnreadStatus('19:c@thread.v2');
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.unreadCount).toBe(2);
    expect(res.value.latestMessageId).toBe('m3');
    expect(res.value.lastReadMessageId).toBe('m1');
  });

  it('propagates a messages fetch failure', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ consumptionhorizons: [] }));
    mockThread.mockResolvedValueOnce({ ok: false, error: { code: 'API_ERROR' } } as never);
    const res = await getUnreadStatus('19:c@thread.v2');
    expect(res.ok).toBe(false);
  });
});

describe('getUnreadConversations', () => {
  it('classifies unread chats and channels', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ conversations: [
      { id: '19:chat@unq.gbl.spaces', properties: { consumptionhorizon: '50;50;50' }, threadProperties: {}, lastMessage: { id: '100', from: '8:orgid:peer', imdisplayname: 'Peer' } },
      { id: '19:chan@thread.tacv2', properties: { consumptionhorizon: '50;50;50' }, threadProperties: { threadType: 'channel', topic: 'General' }, lastMessage: { id: '100', from: '8:orgid:peer', imdisplayname: 'Peer' } },
      { id: '19:read@unq.gbl.spaces', properties: { consumptionhorizon: '200;200;200' }, threadProperties: {}, lastMessage: { id: '100', from: '8:orgid:peer', imdisplayname: 'Peer' } },
    ] }));
    const res = await getUnreadConversations();
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.totalChecked).toBe(3);
    expect(res.value.unreadChats.map(c => c.conversationId)).toContain('19:chat@unq.gbl.spaces');
    expect(res.value.unreadChannels.map(c => c.conversationId)).toContain('19:chan@thread.tacv2');
    expect(res.value.unreadChats.map(c => c.conversationId)).not.toContain('19:read@unq.gbl.spaces');
  });
});

describe('listConversations', () => {
  it('summarises and sorts conversations, honouring a type filter', async () => {
    mockHttp.mockResolvedValue(httpOk({ conversations: [
      { id: '19:chan@thread.tacv2', properties: { favorite: 'true' }, threadProperties: { threadType: 'channel', topic: 'Eng' }, lastMessage: { id: '300', content: '<p>hi</p>', imdisplayname: 'A' } },
      { id: '19:chat@unq.gbl.spaces', properties: {}, threadProperties: {}, lastMessage: { id: '500', content: 'yo', imdisplayname: 'B', from: '8:orgid:peer' } },
    ] }));
    const all = await listConversations({});
    expect(all.ok).toBe(true);
    if (!all.ok) return;
    expect(all.value.total).toBe(2);
    // newest (id 500) first
    expect(all.value.conversations[0].conversationId).toBe('19:chat@unq.gbl.spaces');
    const chan = all.value.conversations.find(c => c.type === 'channel')!;
    expect(chan).toMatchObject({ isFavorite: true, displayName: 'Eng' });

    const channelsOnly = await listConversations({ type: 'channel', limit: 5 });
    expect(channelsOnly.ok).toBe(true);
    if (!channelsOnly.ok) return;
    expect(channelsOnly.value.conversations.every(c => c.type === 'channel')).toBe(true);
  });
});
