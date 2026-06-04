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
  stripHtml: vi.fn((s: string) => `stripped:${s}`),
  extractLinks: vi.fn((content: string) => (content.includes('LINKME') ? [{ url: 'http://x', text: 'x' }] : [])),
  buildMessageLink: vi.fn(() => 'LINK'),
  extractActivityTimestamp: vi.fn((msg: Record<string, unknown>) =>
    (msg.originalarrivaltime === 'NOTS' ? null : (msg.originalarrivaltime ?? null))),
}));

import { httpRequest } from '../utils/http.js';
import { requireMessageAuthWithConfig, getTeamsBaseUrl, getTenantId } from '../utils/auth-guards.js';
import { buildMessageLink } from '../utils/parsers.js';
import { getActivityFeed } from './chatsvc-activity.js';

const mockHttp = vi.mocked(httpRequest);
const mockAuth = vi.mocked(requireMessageAuthWithConfig);
const mockTenant = vi.mocked(getTenantId);
const mockBaseUrl = vi.mocked(getTeamsBaseUrl);
const mockBuildLink = vi.mocked(buildMessageLink);
const AUTH = ok({ auth: { skypeToken: 'sk', authToken: 'at', userMri: '8:orgid:me' }, region: 'uk', baseUrl: 'https://teams.microsoft.com' });
const httpOk = (data: unknown) => ok({ status: 200, headers: new Headers(), data } as never);
const lastUrl = () => mockHttp.mock.calls.at(-1)![0] as string;

beforeEach(() => {
  vi.clearAllMocks();
  mockAuth.mockReturnValue(AUTH as never);
  mockTenant.mockReturnValue('tid');
  mockBaseUrl.mockReturnValue('https://teams.microsoft.com');
});

describe('getActivityFeed', () => {
  it('propagates an auth failure without calling http', async () => {
    mockAuth.mockReturnValue(err(createError(ErrorCode.AUTH_REQUIRED, 'no auth')) as never);
    const res = await getActivityFeed();
    expect(res.ok).toBe(false);
    expect(mockHttp).not.toHaveBeenCalled();
  });

  it('propagates an http failure', async () => {
    mockHttp.mockResolvedValueOnce(err(createError(ErrorCode.NETWORK_ERROR, 'boom')) as never);
    const res = await getActivityFeed();
    expect(res.ok).toBe(false);
  });

  it('uses default page size and no syncState in the URL', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ messages: [] }));
    const res = await getActivityFeed();
    expect(res.ok).toBe(true);
    expect(lastUrl()).toContain('pageSize=50');
    expect(lastUrl()).not.toContain('syncState=');
  });

  it('honours a custom limit and syncState option in the URL', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ messages: [] }));
    await getActivityFeed({ limit: 10, syncState: 'abc def' });
    expect(lastUrl()).toContain('pageSize=10');
    expect(lastUrl()).toContain('syncState=abc%20def');
  });

  it('returns empty activities when messages is not an array', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ messages: 'nope' }));
    const res = await getActivityFeed();
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.activities).toEqual([]);
  });

  it('extracts syncState token from a valid metadata URL', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ messages: [], _metadata: { syncState: 'https://x/feed?syncState=TOKEN123' } }));
    const res = await getActivityFeed();
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.syncState).toBe('TOKEN123');
  });

  it('falls back to the raw metadata syncState when it is not a URL', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ messages: [], _metadata: { syncState: 'invalid' } }));
    const res = await getActivityFeed();
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.syncState).toBe('invalid');
  });

  it('returns undefined syncState when metadata is absent', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ messages: [] }));
    const res = await getActivityFeed();
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.syncState).toBeUndefined();
  });

  it('skips control and system messages and items without id or timestamp', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ messages: [
      {},
      { messagetype: 'Control/Typing' },
      { messagetype: 'ThreadActivity/AddMember' },
      { messagetype: 'ThreadActivity/DeleteMember' },
      { messagetype: 'RichText', id: '', originalarrivaltime: '' },
      { messagetype: 'Text', id: '999', originalarrivaltime: 'NOTS' },
    ] }));
    const res = await getActivityFeed();
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.activities).toEqual([]);
  });

  it('parses, classifies and sorts activities (newest first)', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ messages: [
      { id: '100', messagetype: 'RichText/Html', content: 'itemtype="http://schema.skype.com/Mention"', from: '8:orgid:a', imdisplayname: 'A', originalarrivaltime: '2024-01-03T00:00:00Z', conversationid: '19:abc' },
      { id: '200', messagetype: 'Reaction', content: '', from: '8:orgid:b', originalarrivaltime: '2024-01-02T00:00:00Z', conversationid: '48:notifications', clumpId: '19:real' },
      { id: 'abc', messagetype: 'RichText', threadtopic: 'Topic', from: '8:orgid:c', displayName: 'C', originalarrivaltime: '2024-01-01T00:00:00Z', conversationid: '48:x' },
      { id: '300', messagetype: 'Text', content: 'hi LINKME', from: '8:orgid:d', originalarrivaltime: '2024-01-04T00:00:00Z' },
      { id: '400', messagetype: 'Weird', originalarrivaltime: '2024-01-05T00:00:00Z' },
      { id: '500', messagetype: 'RichText', parentMessageId: 'p1', originalarrivaltime: '2024-01-06T00:00:00Z', conversationId: '19:fromCamel', topic: 'CamelTopic' },
    ] }));
    const res = await getActivityFeed();
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    const a = res.value.activities;
    expect(a).toHaveLength(6);
    expect(a[0].id).toBe('500');
    const byId = Object.fromEntries(a.map(x => [x.id, x]));
    expect(byId['100'].type).toBe('mention');
    expect(byId['100'].activityLink).toBe('LINK');
    expect(byId['100'].sender.displayName).toBe('A');
    expect(byId['200'].type).toBe('reaction');
    expect(byId['200'].conversationId).toBe('19:real');
    expect(byId['200'].activityLink).toBe('LINK');
    expect(byId['abc'].type).toBe('reply');
    expect(byId['abc'].conversationId).toBe('48:x');
    expect(byId['abc'].activityLink).toBeUndefined();
    expect(byId['abc'].sender.displayName).toBe('C');
    expect(byId['abc'].topic).toBe('Topic');
    expect(byId['300'].type).toBe('message');
    expect(byId['300'].links).toEqual([{ url: 'http://x', text: 'x' }]);
    expect(byId['300'].activityLink).toBeUndefined();
    expect(byId['400'].type).toBe('unknown');
    expect(byId['400'].links).toBeUndefined();
    expect(byId['500'].type).toBe('reply');
    expect(byId['500'].conversationId).toBe('19:fromCamel');
    expect(byId['500'].topic).toBe('CamelTopic');
    expect(byId['500'].activityLink).toBe('LINK');
  });

  it('passes tenant undefined to buildMessageLink when tenant id is null', async () => {
    mockTenant.mockReturnValue(null);
    mockHttp.mockResolvedValueOnce(httpOk({ messages: [
      { id: '100', messagetype: 'RichText/Html', content: 'x', from: '8:orgid:a', originalarrivaltime: '2024-01-03T00:00:00Z', conversationid: '19:abc' },
    ] }));
    const res = await getActivityFeed();
    expect(res.ok).toBe(true);
    expect(mockBuildLink).toHaveBeenCalledWith(expect.objectContaining({ tenantId: undefined }));
  });
});
