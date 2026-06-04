import { describe, it, expect, vi, beforeEach } from 'vitest';
import { ok } from '../types/result.js';

vi.mock('../utils/http.js', () => ({ httpRequest: vi.fn() }));
vi.mock('../utils/auth-guards.js', () => ({
  requireMessageAuth: vi.fn(),
  requireMessageAuthWithConfig: vi.fn(),
  getTeamsBaseUrl: vi.fn(() => 'https://teams.microsoft.com'),
  getTenantId: vi.fn(() => 'tenant-123'),
}));
vi.mock('../auth/token-extractor.js', () => ({ getUserDisplayName: vi.fn(() => 'Me') }));

import { httpRequest } from '../utils/http.js';
import {
  requireMessageAuth,
  requireMessageAuthWithConfig,
  getTenantId,
} from '../utils/auth-guards.js';
import { getUserDisplayName } from '../auth/token-extractor.js';
import {
  sendMessage,
  sendNoteToSelf,
  getMessage,
  getThreadMessages,
  editMessage,
  deleteMessage,
  getConversationProperties,
  extractParticipantNames,
  getOneOnOneChatId,
  createGroupChat,
  buildReplyQuoteHtml,
  buildMentionHtml,
  buildMentionsProperty,
  parseContentWithMentionsAndLinks,
} from './chatsvc-messaging.js';

const mockHttp = vi.mocked(httpRequest);
const mockAuthCfg = vi.mocked(requireMessageAuthWithConfig);
const mockAuth = vi.mocked(requireMessageAuth);
const mockName = vi.mocked(getUserDisplayName);
const mockTenant = vi.mocked(getTenantId);

const AUTH_CFG = ok({
  auth: { skypeToken: 'sk', authToken: 'at', userMri: '8:orgid:me' },
  region: 'uk',
  baseUrl: 'https://teams.microsoft.com',
});
const AUTH = ok({ skypeToken: 'sk', authToken: 'at', userMri: '8:orgid:me' });
const FAIL = { ok: false, error: { code: 'AUTH_REQUIRED' } } as never;

const httpOk = (data: unknown = null, headers = new Headers()) =>
  ok({ status: 200, headers, data } as never);
const lastUrl = () => mockHttp.mock.calls.at(-1)![0] as string;
const lastInit = () => mockHttp.mock.calls.at(-1)![1] as { method: string; body?: string; headers: unknown };
const lastBody = () => JSON.parse(lastInit().body!);

const VALID_GUID = 'dde37b63-a4ac-4edb-a7b0-385263022300';
const VALID_GUID_2 = 'aae37b63-a4ac-4edb-a7b0-385263022301';

beforeEach(() => {
  vi.clearAllMocks();
  mockAuthCfg.mockReturnValue(AUTH_CFG as never);
  mockAuth.mockReturnValue(AUTH as never);
  mockName.mockReturnValue('Me');
  mockTenant.mockReturnValue('tenant-123');
});

// ─────────────────────────────────────────────────────────────────────────────
// sendMessage
// ─────────────────────────────────────────────────────────────────────────────

describe('sendMessage', () => {
  it('propagates auth failure without calling http', async () => {
    mockAuthCfg.mockReturnValueOnce(FAIL);
    const res = await sendMessage('19:c@thread.tacv2', 'hi');
    expect(res.ok).toBe(false);
    expect(mockHttp).not.toHaveBeenCalled();
  });

  it('posts a plain channel message and returns the client id + timestamp', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ OriginalArrivalTime: 1780000000000 }));
    const res = await sendMessage('19:c@thread.tacv2', 'hello world');
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.messageId).toMatch(/^\d+$/);
    expect(res.value.timestamp).toBe(1780000000000);
    expect(lastInit().method).toBe('POST');
    const body = lastBody();
    expect(body.messagetype).toBe('RichText/Html');
    expect(body.imdisplayname).toBe('Me');
    expect(body.content).toContain('hello world');
    expect(body.properties).toBeUndefined();
  });

  it('falls back to "User" when no display name is available', async () => {
    mockName.mockReturnValueOnce(null);
    mockHttp.mockResolvedValueOnce(httpOk({}));
    await sendMessage('19:c@thread.tacv2', 'hi');
    expect(lastBody().imdisplayname).toBe('User');
  });

  it('sends rawContentHtml as-is and skips the mention pipeline', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({}));
    await sendMessage('19:c@thread.tacv2', 'ignored', { rawContentHtml: '<p>raw</p>' });
    const body = lastBody();
    expect(body.content).toBe('<p>raw</p>');
    expect(body.properties).toBeUndefined();
  });

  it('includes a mentions property when content has mentions', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({}));
    await sendMessage('19:c@thread.tacv2', 'hey @[Bob](8:orgid:bob)');
    const body = lastBody();
    expect(body.properties.mentions).toBeDefined();
    const mentions = JSON.parse(body.properties.mentions);
    expect(mentions[0]).toMatchObject({ mri: '8:orgid:bob', mentionType: 'person', displayName: 'Bob' });
  });

  it('adds importance for high but not for normal', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({}));
    await sendMessage('19:c@thread.tacv2', 'urgent', { importance: 'high' });
    expect(lastBody().properties.importance).toBe('high');

    mockHttp.mockResolvedValueOnce(httpOk({}));
    await sendMessage('19:c@thread.tacv2', 'meh', { importance: 'normal' });
    expect(lastBody().properties).toBeUndefined();
  });

  it('adds a subject for channel posts', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({}));
    await sendMessage('19:c@thread.tacv2', 'body', { subject: 'Heading' });
    expect(lastBody().subject).toBe('Heading');
  });

  it('uses the threaded URL for a channel reply (native reply chain)', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({}));
    await sendMessage('19:c@thread.tacv2', 'reply', { replyToMessageId: '17800001' });
    expect(lastUrl()).toContain('%3Bmessageid%3D17800001');
    expect(mockHttp).toHaveBeenCalledTimes(1);
  });

  it('embeds a quote block for a chat reply and posts without threading', async () => {
    // first http call: getMessage GET; second: sendMessage POST
    mockHttp
      .mockResolvedValueOnce(httpOk({ id: '17800001', content: '<p>original</p>', from: '8:orgid:bob', imdisplayname: 'Bob', originalarrivaltime: '2026-01-01T00:00:00Z' }))
      .mockResolvedValueOnce(httpOk({}));
    const res = await sendMessage('19:chat@unq.gbl.spaces', 'my reply', { replyToMessageId: '17800001' });
    expect(res.ok).toBe(true);
    expect(mockHttp).toHaveBeenCalledTimes(2);
    const body = lastBody();
    expect(body.content).toContain('schema.skype.com/Reply');
    expect(lastUrl()).not.toContain(';messageid=');
  });

  it('still posts when the quoted message cannot be fetched', async () => {
    mockHttp
      .mockResolvedValueOnce({ ok: false, error: { code: 'NOT_FOUND' } } as never)
      .mockResolvedValueOnce(httpOk({}));
    const res = await sendMessage('19:chat@unq.gbl.spaces', 'my reply', { replyToMessageId: 'bad' });
    expect(res.ok).toBe(true);
    expect(lastBody().content).not.toContain('schema.skype.com/Reply');
  });

  it('returns the http error when the POST fails', async () => {
    mockHttp.mockResolvedValueOnce({ ok: false, error: { code: 'API_ERROR' } } as never);
    const res = await sendMessage('19:c@thread.tacv2', 'hi');
    expect(res.ok).toBe(false);
  });
});

describe('sendNoteToSelf', () => {
  it('sends to the self-chat id', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({}));
    const res = await sendNoteToSelf('note');
    expect(res.ok).toBe(true);
    expect(lastUrl()).toContain('48%3Anotes');
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// getMessage
// ─────────────────────────────────────────────────────────────────────────────

describe('getMessage', () => {
  it('propagates auth failure', async () => {
    mockAuthCfg.mockReturnValueOnce(FAIL);
    const res = await getMessage('19:c@thread.tacv2', '1');
    expect(res.ok).toBe(false);
    expect(mockHttp).not.toHaveBeenCalled();
  });

  it('returns the http error', async () => {
    mockHttp.mockResolvedValueOnce({ ok: false, error: { code: 'API_ERROR' } } as never);
    const res = await getMessage('19:c@thread.tacv2', '1');
    expect(res.ok).toBe(false);
  });

  it('errors when the message has no id', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ content: 'x' }));
    const res = await getMessage('19:c@thread.tacv2', '1');
    expect(res.ok).toBe(false);
    if (res.ok) return;
    expect(res.error.message).toContain('missing ID');
  });

  it('parses a numeric message with a deep link, links, and isFromMe', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({
      id: '17800001',
      content: '<p>see <a href="https://example.com">site</a></p>',
      messagetype: 'RichText/Html',
      from: '8:orgid:me',
      imdisplayname: 'Me',
      originalarrivaltime: '2026-01-30T10:45:00Z',
      clientmessageid: 'cmid',
    }));
    const res = await getMessage('19:c@thread.tacv2', '17800001');
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.id).toBe('17800001');
    expect(res.value.isFromMe).toBe(true);
    expect(res.value.messageLink).toBeDefined();
    expect(res.value.links!.length).toBeGreaterThan(0);
    expect(res.value.when).toBeDefined();
    expect(res.value.clientMessageId).toBe('cmid');
  });

  it('flags a thread reply and uses composetime / non-numeric id fallback', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({
      id: 'abc-non-numeric',
      messagetype: 'RichText/Html',
      from: '8:orgid:other',
      displayName: 'Other',
      composetime: '2026-01-01T00:00:00Z',
      rootMessageId: 'root-1',
    }));
    const res = await getMessage('19:c@thread.tacv2', 'abc-non-numeric');
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.isThreadReply).toBe(true);
    expect(res.value.threadRootId).toBe('root-1');
    expect(res.value.messageLink).toBeUndefined();
    expect(res.value.isFromMe).toBe(false);
    expect(res.value.sender.displayName).toBe('Other');
  });

  it('falls back to id from originalarrivaltime when id absent', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({
      originalarrivaltime: '17800002',
      content: '',
    }));
    const res = await getMessage('19:c@thread.tacv2', 'x');
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.id).toBe('17800002');
    expect(res.value.contentType).toBe('Text');
    expect(res.value.links).toBeUndefined();
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// getThreadMessages
// ─────────────────────────────────────────────────────────────────────────────

describe('getThreadMessages', () => {
  it('propagates auth failure', async () => {
    mockAuthCfg.mockReturnValueOnce(FAIL);
    const res = await getThreadMessages('19:c@thread.tacv2');
    expect(res.ok).toBe(false);
  });

  it('returns the http error', async () => {
    mockHttp.mockResolvedValueOnce({ ok: false, error: { code: 'API_ERROR' } } as never);
    const res = await getThreadMessages('19:c@thread.tacv2');
    expect(res.ok).toBe(false);
  });

  it('returns empty when messages is not an array', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ messages: 'nope' }));
    const res = await getThreadMessages('19:c@thread.tacv2');
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.messages).toEqual([]);
  });

  it('skips control/system/deleted/idless messages and parses real ones (desc default)', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ messages: [
      { messagetype: 'Control/Typing', id: '1' },
      { messagetype: 'ThreadActivity/AddMember', id: '2' },
      { id: '3', properties: { deletetime: 123 }, messagetype: 'RichText/Html' },
      { messagetype: 'RichText/Html', content: 'x' }, // no id
      { id: '17800001', messagetype: 'RichText/Html', content: '<p>older</p>', from: '8:orgid:me', imdisplayname: 'Me', originalarrivaltime: '2026-01-01T00:00:00Z' },
      { id: '17800002', messagetype: 'RichText/Html', content: '<p>newer <a href="https://x.io">x</a></p>', from: '8:orgid:bob', imdisplayname: 'Bob', originalarrivaltime: '2026-02-01T00:00:00Z', rootMessageId: 'root', clientmessageid: 'cm' },
      { content: 'no type' }, // no messagetype -> skipped
    ] }));
    const res = await getThreadMessages('19:c@thread.tacv2');
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.messages.map(m => m.id)).toEqual(['17800002', '17800001']);
    const newer = res.value.messages[0];
    expect(newer.isThreadReply).toBe(true);
    expect(newer.threadRootId).toBe('root');
    expect(newer.isFromMe).toBe(false);
    expect(newer.links!.length).toBeGreaterThan(0);
    expect(newer.messageLink).toBeDefined();
  });

  it('honours asc order, limit, and startTime in the URL', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ messages: [
      { id: '17800001', messagetype: 'RichText/Html', content: 'a', originalarrivaltime: '2026-01-01T00:00:00Z' },
      { id: '17800002', messagetype: 'RichText/Html', content: 'b', originalarrivaltime: '2026-02-01T00:00:00Z' },
    ] }));
    const res = await getThreadMessages('19:c@thread.tacv2', { limit: 10, startTime: 99, order: 'asc' });
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.messages.map(m => m.id)).toEqual(['17800001', '17800002']);
    expect(lastUrl()).toContain('pageSize=10');
    expect(lastUrl()).toContain('startTime=99');
  });

  it('uses id-derived timestamp when no time fields and skips deep link for non-numeric id', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ messages: [
      { id: 'non-numeric', messagetype: 'RichText/Html', content: 'x' },
    ] }));
    const res = await getThreadMessages('19:c@thread.tacv2');
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.messages[0].messageLink).toBeUndefined();
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// editMessage / deleteMessage
// ─────────────────────────────────────────────────────────────────────────────

describe('editMessage', () => {
  it('propagates auth failure', async () => {
    mockAuthCfg.mockReturnValueOnce(FAIL);
    const res = await editMessage('19:c@thread.tacv2', '1', 'x');
    expect(res.ok).toBe(false);
  });

  it('PUTs new content without mentions', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({}));
    const res = await editMessage('19:c@thread.tacv2', '17800001', 'updated text');
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value).toEqual({ messageId: '17800001', conversationId: '19:c@thread.tacv2' });
    expect(lastInit().method).toBe('PUT');
    const body = lastBody();
    expect(body.id).toBe('17800001');
    expect(body.properties).toBeUndefined();
  });

  it('includes mentions when present', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({}));
    await editMessage('19:c@thread.tacv2', '1', 'hi @[Bob](8:orgid:bob)');
    expect(lastBody().properties.mentions).toBeDefined();
  });

  it('returns the http error', async () => {
    mockHttp.mockResolvedValueOnce({ ok: false, error: { code: 'API_ERROR' } } as never);
    const res = await editMessage('19:c@thread.tacv2', '1', 'x');
    expect(res.ok).toBe(false);
  });
});

describe('deleteMessage', () => {
  it('propagates auth failure', async () => {
    mockAuthCfg.mockReturnValueOnce(FAIL);
    const res = await deleteMessage('19:c@thread.tacv2', '1');
    expect(res.ok).toBe(false);
  });

  it('DELETEs and returns ids', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({}));
    const res = await deleteMessage('19:c@thread.tacv2', '17800001');
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value).toEqual({ messageId: '17800001', conversationId: '19:c@thread.tacv2' });
    expect(lastInit().method).toBe('DELETE');
  });

  it('returns the http error', async () => {
    mockHttp.mockResolvedValueOnce({ ok: false, error: { code: 'API_ERROR' } } as never);
    const res = await deleteMessage('19:c@thread.tacv2', '1');
    expect(res.ok).toBe(false);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// getConversationProperties
// ─────────────────────────────────────────────────────────────────────────────

describe('getConversationProperties', () => {
  it('propagates auth failure', async () => {
    mockAuthCfg.mockReturnValueOnce(FAIL);
    const res = await getConversationProperties('19:c@thread.tacv2');
    expect(res.ok).toBe(false);
  });

  it('returns the http error', async () => {
    mockHttp.mockResolvedValueOnce({ ok: false, error: { code: 'API_ERROR' } } as never);
    const res = await getConversationProperties('19:c@thread.tacv2');
    expect(res.ok).toBe(false);
  });

  it('uses topicThreadTopic and Meeting product type', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ threadProperties: { topicThreadTopic: 'T1', productThreadType: 'Meeting' } }));
    const res = await getConversationProperties('19:x');
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value).toEqual({ displayName: 'T1', conversationType: 'Meeting' });
  });

  it('uses topic and Channel product type', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ threadProperties: { topic: 'T2', productThreadType: 'TeamThreadChannel' } }));
    const res = await getConversationProperties('19:x');
    if (!res.ok) return;
    expect(res.value).toEqual({ displayName: 'T2', conversationType: 'Channel' });
  });

  it('uses spaceThreadTopic and Chat product type', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ threadProperties: { spaceThreadTopic: 'T3', productThreadType: 'Chat' } }));
    const res = await getConversationProperties('19:x');
    if (!res.ok) return;
    expect(res.value).toEqual({ displayName: 'T3', conversationType: 'Chat' });
  });

  it('uses threadtopic and TeamsTeam product type', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ threadProperties: { threadtopic: 'T4', productThreadType: 'TeamsTeam' } }));
    const res = await getConversationProperties('19:x');
    if (!res.ok) return;
    expect(res.value).toEqual({ displayName: 'T4', conversationType: 'Channel' });
  });

  it('builds a display name from up to 3 members and detects meeting_ id', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ members: [
      { mri: '8:orgid:me', friendlyName: 'Me' },
      { mri: '8:orgid:a', friendlyName: 'Alice' },
      { id: '8:orgid:b', displayName: 'Bob' },
      { mri: '8:orgid:c', name: 'Carol' },
    ] }));
    const res = await getConversationProperties('19:meeting_abc@thread.v2');
    if (!res.ok) return;
    expect(res.value.displayName).toBe('Alice, Bob, Carol');
    expect(res.value.conversationType).toBe('Meeting');
  });

  it('summarises more than 3 members with a +N more and detects channel via groupId', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ threadProperties: { groupId: 'g1' }, members: [
      { mri: '8:orgid:a', friendlyName: 'Alice' },
      { mri: '8:orgid:b', friendlyName: 'Bob' },
      { mri: '8:orgid:c', friendlyName: 'Carol' },
      { mri: '8:orgid:d', friendlyName: 'Dave' },
    ] }));
    const res = await getConversationProperties('19:x');
    if (!res.ok) return;
    expect(res.value.displayName).toBe('Alice, Bob, Carol + 1 more');
    expect(res.value.conversationType).toBe('Channel');
  });

  it('detects chat from @thread.v2 id when no other signal', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ members: [{ mri: '8:orgid:me', friendlyName: 'Me' }] }));
    const res = await getConversationProperties('19:abc@thread.v2');
    if (!res.ok) return;
    expect(res.value.displayName).toBeUndefined();
    expect(res.value.conversationType).toBe('Chat');
  });

  it('detects chat from an 8: prefixed id', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({}));
    const res = await getConversationProperties('8:orgid:peer');
    if (!res.ok) return;
    expect(res.value.conversationType).toBe('Chat');
  });

  it('leaves type undefined when nothing matches', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({}));
    const res = await getConversationProperties('19:weird@unknown');
    if (!res.ok) return;
    expect(res.value.conversationType).toBeUndefined();
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// extractParticipantNames
// ─────────────────────────────────────────────────────────────────────────────

describe('extractParticipantNames', () => {
  it('returns undefined (ok) when not authenticated', async () => {
    mockAuthCfg.mockReturnValueOnce(FAIL);
    const res = await extractParticipantNames('19:c@thread.tacv2');
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value).toBeUndefined();
  });

  it('returns undefined when the request fails', async () => {
    mockHttp.mockResolvedValueOnce({ ok: false, error: { code: 'API_ERROR' } } as never);
    const res = await extractParticipantNames('19:c@thread.tacv2');
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value).toBeUndefined();
  });

  it('returns undefined when there are no messages', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ messages: [] }));
    const res = await extractParticipantNames('19:c@thread.tacv2');
    if (!res.ok) return;
    expect(res.value).toBeUndefined();
  });

  it('returns undefined when only self or nameless senders exist', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ messages: [
      { from: '8:orgid:me', imdisplayname: 'Me' },
      { from: '8:orgid:x' },
    ] }));
    const res = await extractParticipantNames('19:c@thread.tacv2');
    if (!res.ok) return;
    expect(res.value).toBeUndefined();
  });

  it('joins up to 3 unique names', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ messages: [
      { from: '8:orgid:a', imdisplayname: 'Alice' },
      { from: '8:orgid:a', imdisplayname: 'Alice' },
      { from: '8:orgid:b', imdisplayname: 'Bob' },
    ] }));
    const res = await extractParticipantNames('19:c@thread.tacv2');
    if (!res.ok) return;
    expect(res.value).toBe('Alice, Bob');
  });

  it('summarises more than 3 names with +N more', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ messages: [
      { from: '8:orgid:a', imdisplayname: 'Alice' },
      { from: '8:orgid:b', imdisplayname: 'Bob' },
      { from: '8:orgid:c', imdisplayname: 'Carol' },
      { from: '8:orgid:d', imdisplayname: 'Dave' },
    ] }));
    const res = await extractParticipantNames('19:c@thread.tacv2');
    if (!res.ok) return;
    expect(res.value).toBe('Alice, Bob, Carol + 1 more');
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// getOneOnOneChatId
// ─────────────────────────────────────────────────────────────────────────────

describe('getOneOnOneChatId', () => {
  it('propagates auth failure', () => {
    mockAuth.mockReturnValueOnce(FAIL);
    const res = getOneOnOneChatId(VALID_GUID);
    expect(res.ok).toBe(false);
  });

  it('errors when the current user MRI has no extractable id', () => {
    mockAuth.mockReturnValueOnce(ok({ skypeToken: 'sk', authToken: 'at', userMri: 'garbage' }) as never);
    const res = getOneOnOneChatId(VALID_GUID);
    expect(res.ok).toBe(false);
    if (res.ok) return;
    expect(res.error.code).toBe('AUTH_REQUIRED');
  });

  it('errors on an invalid other identifier', () => {
    mockAuth.mockReturnValueOnce(ok({ skypeToken: 'sk', authToken: 'at', userMri: `8:orgid:${VALID_GUID}` }) as never);
    const res = getOneOnOneChatId('not-an-id');
    expect(res.ok).toBe(false);
    if (res.ok) return;
    expect(res.error.code).toBe('INVALID_INPUT');
  });

  it('builds the sorted 1:1 conversation id', () => {
    mockAuth.mockReturnValueOnce(ok({ skypeToken: 'sk', authToken: 'at', userMri: `8:orgid:${VALID_GUID}` }) as never);
    const res = getOneOnOneChatId(VALID_GUID_2);
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.conversationId).toBe(`19:${VALID_GUID_2}_${VALID_GUID}@unq.gbl.spaces`);
    expect(res.value.currentUserId).toBe(VALID_GUID);
    expect(res.value.otherUserId).toBe(VALID_GUID_2);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// createGroupChat
// ─────────────────────────────────────────────────────────────────────────────

describe('createGroupChat', () => {
  it('propagates auth failure', async () => {
    mockAuthCfg.mockReturnValueOnce(FAIL);
    const res = await createGroupChat([VALID_GUID, VALID_GUID_2]);
    expect(res.ok).toBe(false);
  });

  it('rejects fewer than 2 members', async () => {
    const res = await createGroupChat([VALID_GUID]);
    expect(res.ok).toBe(false);
    if (res.ok) return;
    expect(res.error.code).toBe('INVALID_INPUT');
    expect(mockHttp).not.toHaveBeenCalled();
  });

  it('rejects duplicate members', async () => {
    const res = await createGroupChat([VALID_GUID, VALID_GUID]);
    expect(res.ok).toBe(false);
    if (res.ok) return;
    expect(res.error.message).toContain('Duplicate');
  });

  it('rejects more than 250 members', async () => {
    const many = Array.from({ length: 251 }, (_, i) => `8:orgid:${i}`);
    const res = await createGroupChat(many);
    expect(res.ok).toBe(false);
    if (res.ok) return;
    expect(res.error.message).toContain('250');
  });

  it('rejects an invalid member identifier', async () => {
    const res = await createGroupChat([VALID_GUID, 'not-valid']);
    expect(res.ok).toBe(false);
    if (res.ok) return;
    expect(res.error.code).toBe('INVALID_INPUT');
    expect(mockHttp).not.toHaveBeenCalled();
  });

  it('creates the chat from threadResource.id and includes the current user MRI + topic', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ threadResource: { id: '19:new@thread.v2' } }));
    const res = await createGroupChat([`8:orgid:${VALID_GUID}`, VALID_GUID_2], 'Team Chat');
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.conversationId).toBe('19:new@thread.v2');
    expect(res.value.topic).toBe('Team Chat');
    expect(res.value.members[0]).toBe('8:orgid:me');
    expect(res.value.members).toContain(`8:orgid:${VALID_GUID}`);
    expect(res.value.members).toContain(`8:orgid:${VALID_GUID_2}`);
    const body = lastBody();
    expect(body.properties.threadType).toBe('chat');
    expect(body.properties.topic).toBe('Team Chat');
  });

  it('uses responseData.id then threadId fallbacks', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ id: '19:byid@thread.v2' }));
    const a = await createGroupChat([VALID_GUID, VALID_GUID_2]);
    if (!a.ok) return;
    expect(a.value.conversationId).toBe('19:byid@thread.v2');

    mockHttp.mockResolvedValueOnce(httpOk({ threadId: '19:bythreadid@thread.v2' }));
    const b = await createGroupChat([VALID_GUID, VALID_GUID_2]);
    if (!b.ok) return;
    expect(b.value.conversationId).toBe('19:bythreadid@thread.v2');
  });

  it('extracts the id from the Location header', async () => {
    const headers = new Headers({ location: 'https://x/threads/19:fromheader@thread.v2' });
    mockHttp.mockResolvedValueOnce(httpOk({}, headers));
    const res = await createGroupChat([VALID_GUID, VALID_GUID_2]);
    if (!res.ok) return;
    expect(res.value.conversationId).toBe('19:fromheader@thread.v2');
  });

  it('returns a note when no id can be found', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({}));
    const res = await createGroupChat([VALID_GUID, VALID_GUID_2]);
    if (!res.ok) return;
    expect(res.value.note).toContain('did not return');
    expect(res.value.conversationId).toContain('created');
  });

  it('returns the http error', async () => {
    mockHttp.mockResolvedValueOnce({ ok: false, error: { code: 'API_ERROR' } } as never);
    const res = await createGroupChat([VALID_GUID, VALID_GUID_2]);
    expect(res.ok).toBe(false);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// Pure exported helpers
// ─────────────────────────────────────────────────────────────────────────────

describe('buildReplyQuoteHtml', () => {
  it('normalises the sender MRI and escapes content', () => {
    const html = buildReplyQuoteHtml({
      id: '17800001',
      content: '<b>hello</b>',
      contentType: 'text',
      sender: { mri: '8:orgid:bob/extra"', displayName: 'Bob & Co' },
      timestamp: 't',
      conversationId: 'c',
    });
    expect(html).toContain('schema.skype.com/Reply');
    expect(html).toContain('itemid="17800001"');
    expect(html).toContain('Bob &amp; Co');
    expect(html).toContain('8:orgid:bob');
  });

  it('falls back to the raw mri and Unknown name', () => {
    const html = buildReplyQuoteHtml({
      id: 'x',
      content: 'hi',
      contentType: 'text',
      sender: { mri: 'weirdmri' },
      timestamp: 't',
      conversationId: 'c',
    });
    expect(html).toContain('itemid="weirdmri"');
    expect(html).toContain('Unknown');
  });
});

describe('buildMentionHtml', () => {
  it('wraps person mentions in a readonly element', () => {
    const html = buildMentionHtml('Bob', 0, '8:orgid:bob');
    expect(html).toContain('<readonly');
    expect(html).toContain('itemid="0"');
  });

  it('uses a span-only format for tag mentions', () => {
    const html = buildMentionHtml('Eng', 1, 'tag:abc');
    expect(html).not.toContain('<readonly');
    expect(html).toContain('<span');
  });
});

describe('buildMentionsProperty', () => {
  it('encodes person and tag mentions with the right type and stripped tag prefix', () => {
    const json = buildMentionsProperty([
      { mri: '8:orgid:bob', displayName: 'Bob' },
      { mri: 'tag:abc', displayName: 'Eng' },
    ]);
    const arr = JSON.parse(json);
    expect(arr[0]).toMatchObject({ mri: '8:orgid:bob', mentionType: 'person', itemid: '0' });
    expect(arr[1]).toMatchObject({ mri: 'abc', mentionType: 'tag', itemid: '1' });
  });
});

describe('parseContentWithMentionsAndLinks', () => {
  it('runs plain markdown when there are no mentions or links', () => {
    const { html, mentions } = parseContentWithMentionsAndLinks('**bold**');
    expect(mentions).toEqual([]);
    expect(html).toContain('<');
  });

  it('extracts mentions and substitutes their html', () => {
    const { html, mentions } = parseContentWithMentionsAndLinks('hi @[Bob](8:orgid:bob)');
    expect(mentions).toEqual([{ mri: '8:orgid:bob', displayName: 'Bob' }]);
    expect(html).toContain('schema.skype.com/Mention');
  });

  it('renders links as anchors and escapes quotes in the url', () => {
    const { html, mentions } = parseContentWithMentionsAndLinks('see [site](https://e.com/a"b)');
    expect(mentions).toEqual([]);
    expect(html).toContain('<a href="https://e.com/a&quot;b">');
    expect(html).toContain('site');
  });

  it('handles a mix of mentions and links in source order', () => {
    const { html, mentions } = parseContentWithMentionsAndLinks('@[A](8:orgid:a) then [x](https://e.com) and @[B](tag:t)');
    expect(mentions).toEqual([
      { mri: '8:orgid:a', displayName: 'A' },
      { mri: 'tag:t', displayName: 'B' },
    ]);
    expect(html).toContain('<a href="https://e.com">');
  });
});
