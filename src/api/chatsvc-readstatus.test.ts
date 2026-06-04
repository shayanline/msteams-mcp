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

const AUTH_ERR = { ok: false, error: { code: 'AUTH_REQUIRED' } } as never;
const httpErr = { ok: false, error: { code: 'API_ERROR' } } as never;

describe('getConsumptionHorizon edge cases', () => {
  it('returns the auth error without calling http', async () => {
    mockAuth.mockReturnValueOnce(AUTH_ERR);
    const res = await getConsumptionHorizon('19:c@thread.v2');
    expect(res.ok).toBe(false);
    expect(mockHttp).not.toHaveBeenCalled();
  });

  it('propagates an http failure', async () => {
    mockHttp.mockResolvedValueOnce(httpErr);
    expect((await getConsumptionHorizon('19:c@thread.v2')).ok).toBe(false);
  });

  it('defaults to no horizons and matches a partial id with a short horizon', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ version: 'v2' }));
    const empty = await getConsumptionHorizon('19:c@thread.v2');
    expect(empty.ok).toBe(true);
    if (!empty.ok) return;
    expect(empty.value.consumptionHorizons).toEqual([]);
    expect(empty.value.lastReadMessageId).toBeUndefined();

    // id includes the user MRI (but isn't equal) and a horizon with fewer than 3 parts
    mockHttp.mockResolvedValueOnce(httpOk({ consumptionhorizons: [
      { id: '8:orgid:me;device', consumptionhorizon: '100' },
    ] }));
    const partial = await getConsumptionHorizon('19:c@thread.v2');
    expect(partial.ok).toBe(true);
    if (!partial.ok) return;
    expect(partial.value.lastReadMessageId).toBeUndefined();
    expect(partial.value.lastReadTimestamp).toBeUndefined();
  });
});

describe('markAsRead / markUnread error paths', () => {
  it('markAsRead returns the auth error', async () => {
    mockAuth.mockReturnValueOnce(AUTH_ERR);
    expect((await markAsRead('19:c@thread.v2', '1')).ok).toBe(false);
    expect(mockHttp).not.toHaveBeenCalled();
  });

  it('markAsRead propagates an http failure', async () => {
    mockHttp.mockResolvedValueOnce(httpErr);
    expect((await markAsRead('19:c@thread.v2', '1')).ok).toBe(false);
  });

  it('markUnread returns the auth error', async () => {
    mockAuth.mockReturnValueOnce(AUTH_ERR);
    expect((await markUnread('19:c@thread.v2', '1')).ok).toBe(false);
    expect(mockHttp).not.toHaveBeenCalled();
  });

  it('markUnread propagates an http failure', async () => {
    mockHttp.mockResolvedValueOnce(httpErr);
    expect((await markUnread('19:c@thread.v2', '1')).ok).toBe(false);
  });
});

describe('getUnreadStatus without a read horizon', () => {
  it('falls back when the horizon lookup fails and skips own messages for latest', async () => {
    mockHttp.mockResolvedValueOnce(httpErr); // getConsumptionHorizon fails -> lastReadId undefined
    mockThread.mockResolvedValueOnce(ok({ conversationId: '19:c@thread.v2', messages: [
      { id: 'm3', isFromMe: true },   // own message: not latest, not counted
      { id: 'm2', isFromMe: false },  // counted, becomes latest
      { id: 'm1', isFromMe: false },  // counted
    ] }) as never);
    const res = await getUnreadStatus('19:c@thread.v2');
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.unreadCount).toBe(2);
    expect(res.value.latestMessageId).toBe('m2');
    expect(res.value.lastReadMessageId).toBeUndefined();
  });
});

describe('getUnreadConversations edge cases', () => {
  it('returns the auth error without calling http', async () => {
    mockAuth.mockReturnValueOnce(AUTH_ERR);
    expect((await getUnreadConversations()).ok).toBe(false);
    expect(mockHttp).not.toHaveBeenCalled();
  });

  it('propagates an http failure', async () => {
    mockHttp.mockResolvedValueOnce(httpErr);
    expect((await getUnreadConversations()).ok).toBe(false);
  });

  it('defaults to no conversations when the payload has no data', async () => {
    mockHttp.mockResolvedValueOnce(httpOk(undefined));
    const res = await getUnreadConversations();
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.totalChecked).toBe(0);
  });

  it('classifies every branch of conversation state', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ conversations: [
      // no lastMessage id -> skipped
      { id: '19:a@unq.gbl.spaces', properties: {}, threadProperties: {}, lastMessage: {} },
      // non-numeric last message id -> skipped
      { id: '19:b@unq.gbl.spaces', properties: {}, threadProperties: {}, lastMessage: { id: 'xyz' } },
      // never-opened, last message from me -> skipped
      { id: '19:c@unq.gbl.spaces', properties: {}, threadProperties: {}, lastMessage: { id: '100', from: '8:orgid:me', imdisplayname: 'Me' } },
      // never-opened channel -> skipped
      { id: '19:d@thread.tacv2', properties: {}, threadProperties: { threadType: 'channel' }, lastMessage: { id: '100', from: '8:orgid:peer', imdisplayname: 'P' } },
      // never-opened DM with no sender -> pushed as unread chat (readUpTo 0)
      { id: '19:e@unq.gbl.spaces', properties: {}, threadProperties: {}, lastMessage: { id: '100', imdisplayname: 'P' } },
      // malformed horizon -> skipped
      { id: '19:f@unq.gbl.spaces', properties: { consumptionhorizon: 'abc;1;2' }, threadProperties: {}, lastMessage: { id: '100', from: '8:orgid:peer' } },
      // already read -> skipped
      { id: '19:g@unq.gbl.spaces', properties: { consumptionhorizon: '200;200;200' }, threadProperties: {}, lastMessage: { id: '100', from: '8:orgid:peer' } },
      // self reply within the read window -> skipped
      { id: '19:h@unq.gbl.spaces', properties: { consumptionhorizon: '99500;99500;99500' }, threadProperties: {}, lastMessage: { id: '100000', from: '8:orgid:me' } },
      // self reply in a channel -> skipped
      { id: '19:i@thread.tacv2', properties: { consumptionhorizon: '1;1;1' }, threadProperties: { threadType: 'channel' }, lastMessage: { id: '100000', from: '8:orgid:me' } },
      // self reply with a large gap -> still unread chat
      { id: '19:j@unq.gbl.spaces', properties: { consumptionhorizon: '1;1;1' }, threadProperties: {}, lastMessage: { id: '100000', from: '8:orgid:me', imdisplayname: 'Me' } },
      // normal unread chat
      { id: '19:k@unq.gbl.spaces', properties: { consumptionhorizon: '50;50;50' }, threadProperties: { topic: 'Topic K' }, lastMessage: { id: '100', from: '8:orgid:peer', imdisplayname: 'P' } },
      // normal unread channel
      { id: '19:l@thread.tacv2', properties: { consumptionhorizon: '50;50;50' }, threadProperties: { threadType: 'channel' }, lastMessage: { id: '100', from: '8:orgid:peer', imdisplayname: 'P' } },
    ] }));
    const res = await getUnreadConversations();
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    const chatIds = res.value.unreadChats.map(c => c.conversationId);
    const chanIds = res.value.unreadChannels.map(c => c.conversationId);
    expect(chatIds).toEqual(expect.arrayContaining(['19:e@unq.gbl.spaces', '19:j@unq.gbl.spaces', '19:k@unq.gbl.spaces']));
    expect(chanIds).toContain('19:l@thread.tacv2');
    expect(chatIds).not.toContain('19:c@unq.gbl.spaces');
    expect(chatIds).not.toContain('19:g@unq.gbl.spaces');
    expect(chatIds).not.toContain('19:h@unq.gbl.spaces');
    expect(res.value.totalChecked).toBe(12);
  });
});

describe('listConversations edge cases', () => {
  it('returns the auth error without calling http', async () => {
    mockAuth.mockReturnValueOnce(AUTH_ERR);
    expect((await listConversations()).ok).toBe(false);
    expect(mockHttp).not.toHaveBeenCalled();
  });

  it('propagates an http failure', async () => {
    mockHttp.mockResolvedValueOnce(httpErr);
    expect((await listConversations()).ok).toBe(false);
  });

  it('skips entries with no id and classifies meetings, previews and unread state', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ conversations: [
      // no id -> skipped
      { properties: {}, threadProperties: {}, lastMessage: { id: '1' } },
      // meeting via threadType
      { id: '19:meet1@thread.v2', properties: {}, threadProperties: { threadType: 'meeting', spaceThreadTopic: 'Weekly' }, lastMessage: { id: '400', from: '8:orgid:peer', content: '<p>hello team</p>', imdisplayname: 'P' } },
      // meeting via id + topic mentioning meeting
      { id: '19:meet2@thread.v2', properties: { consumptionhorizon: '10;10;10' }, threadProperties: { spaceThreadTopic: 'Project Meeting' }, lastMessage: { id: '200', from: '8:orgid:peer', imdisplayname: 'P' } },
      // read chat (lastMsgTime <= readUpTo) -> not unread
      { id: '19:rd@unq.gbl.spaces', properties: { consumptionhorizon: '900;900;900' }, threadProperties: {}, lastMessage: { id: '500', from: '8:orgid:peer', imdisplayname: 'P' } },
      // chat with a non-numeric last message id -> lastMessageTime undefined
      { id: '19:nan@unq.gbl.spaces', properties: {}, threadProperties: {}, lastMessage: { id: 'abc', imdisplayname: 'P' } },
    ] }));
    const res = await listConversations({});
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    const byId = Object.fromEntries(res.value.conversations.map(c => [c.conversationId, c]));
    expect(byId['19:meet1@thread.v2'].type).toBe('meeting');
    expect(byId['19:meet1@thread.v2'].lastMessagePreview).toBe('hello team');
    expect(byId['19:meet1@thread.v2'].isUnread).toBe(true);
    expect(byId['19:meet2@thread.v2'].type).toBe('meeting');
    expect(byId['19:rd@unq.gbl.spaces'].isUnread).toBe(false);
    expect(byId['19:nan@unq.gbl.spaces'].lastMessageTime).toBeUndefined();
    expect(byId['19:nan@unq.gbl.spaces'].lastMessagePreview).toBeUndefined();
    expect(res.value.total).toBe(4); // the id-less entry was skipped
  });
});
