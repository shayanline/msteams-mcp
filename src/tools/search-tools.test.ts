/**
 * Unit tests for search tools (real handlers).
 */
import { describe, it, expect, vi, beforeEach } from 'vitest';
import { ok, err } from '../types/result.js';
import { createError, ErrorCode } from '../types/errors.js';
import './index.js';

vi.mock('../api/substrate-api.js', () => ({
  searchMessages: vi.fn(),
  searchEmails: vi.fn(),
  searchChannels: vi.fn(),
}));
vi.mock('../api/chatsvc-api.js', () => ({
  getThreadMessages: vi.fn(),
  getConsumptionHorizon: vi.fn(),
  markAsRead: vi.fn(),
}));

import { searchMessages, searchEmails, searchChannels } from '../api/substrate-api.js';
import { getThreadMessages, getConsumptionHorizon, markAsRead } from '../api/chatsvc-api.js';
import { searchTool, getThreadTool, findChannelTool, searchEmailTool } from './search-tools.js';

const ctx = { server: {} } as never;
const anErr = err(createError(ErrorCode.API_ERROR, 'boom'));

const pagination = (hasMore: boolean) => ({ from: 0, size: 25, returned: 2, total: 10, hasMore });

beforeEach(() => vi.clearAllMocks());

describe('searchTool', () => {
  it('returns results with nextFrom when hasMore', async () => {
    vi.mocked(searchMessages).mockResolvedValue(ok({ results: [{}, {}], pagination: pagination(true) }) as never);
    const res = await searchTool.handler({ query: 'hi', maxResults: 25, from: 0, size: 25 }, ctx);
    expect(res.success).toBe(true);
    if (res.success) expect(res.data.pagination).toMatchObject({ hasMore: true, nextFrom: 2 });
  });

  it('omits nextFrom when no more results', async () => {
    vi.mocked(searchMessages).mockResolvedValue(ok({ results: [{}], pagination: pagination(false) }) as never);
    const res = await searchTool.handler({ query: 'hi', maxResults: 25, from: 0, size: 25 }, ctx);
    expect(res.success).toBe(true);
    if (res.success) expect((res.data.pagination as Record<string, unknown>).nextFrom).toBeUndefined();
  });

  it('propagates errors', async () => {
    vi.mocked(searchMessages).mockResolvedValue(anErr as never);
    const res = await searchTool.handler({ query: 'hi', maxResults: 25, from: 0, size: 25 }, ctx);
    expect(res.success).toBe(false);
  });
});

describe('searchEmailTool', () => {
  it('returns results and includes calendar filtered note', async () => {
    vi.mocked(searchEmails).mockResolvedValue(
      ok({ results: [{}, {}], pagination: pagination(true), filteredCount: 3 }) as never
    );
    const res = await searchEmailTool.handler({ query: 'hi', maxResults: 25, from: 0, size: 25 }, ctx);
    expect(res.success).toBe(true);
    if (res.success) expect(res.data.calendarResponsesFiltered).toBe(3);
  });

  it('omits note when no filtered results', async () => {
    vi.mocked(searchEmails).mockResolvedValue(ok({ results: [{}], pagination: pagination(false) }) as never);
    const res = await searchEmailTool.handler({ query: 'hi', maxResults: 25, from: 0, size: 25 }, ctx);
    expect(res.success).toBe(true);
    if (res.success) expect(res.data.calendarResponsesFiltered).toBeUndefined();
  });

  it('propagates errors', async () => {
    vi.mocked(searchEmails).mockResolvedValue(anErr as never);
    const res = await searchEmailTool.handler({ query: 'x', maxResults: 25, from: 0, size: 25 }, ctx);
    expect(res.success).toBe(false);
  });
});

describe('findChannelTool', () => {
  it('returns channels on success', async () => {
    vi.mocked(searchChannels).mockResolvedValue(ok({ returned: 1, results: [{ id: 'ch1' }] }) as never);
    const res = await findChannelTool.handler({ query: 'eng', limit: 10 }, ctx);
    expect(res.success).toBe(true);
    if (res.success) expect(res.data).toEqual({ query: 'eng', count: 1, channels: [{ id: 'ch1' }] });
  });

  it('propagates errors', async () => {
    vi.mocked(searchChannels).mockResolvedValue(anErr as never);
    const res = await findChannelTool.handler({ query: 'eng', limit: 10 }, ctx);
    expect(res.success).toBe(false);
  });
});

describe('getThreadTool', () => {
  const baseInput = { conversationId: 'c1', limit: 50, markRead: false, order: 'desc' as const };

  it('rejects invalid since timestamp', async () => {
    const res = await getThreadTool.handler({ ...baseInput, since: 'not-a-date' }, ctx);
    expect(res.success).toBe(false);
    if (!res.success) expect(res.error.code).toBe(ErrorCode.INVALID_INPUT);
  });

  it('accepts a valid since timestamp', async () => {
    vi.mocked(getThreadMessages).mockResolvedValue(ok({ conversationId: 'c1', messages: [] }) as never);
    vi.mocked(getConsumptionHorizon).mockResolvedValue(err(createError(ErrorCode.API_ERROR, 'x')) as never);
    const res = await getThreadTool.handler({ ...baseInput, since: '2026-02-26T00:00:00Z' }, ctx);
    expect(res.success).toBe(true);
  });

  it('propagates thread fetch errors', async () => {
    vi.mocked(getThreadMessages).mockResolvedValue(anErr as never);
    const res = await getThreadTool.handler(baseInput, ctx);
    expect(res.success).toBe(false);
  });

  it('computes unread count when last-read message is in window', async () => {
    const messages = [
      { id: 'm1', timestamp: '2026-01-01T00:00:00Z', isFromMe: false },
      { id: 'm2', timestamp: '2026-01-02T00:00:00Z', isFromMe: false },
    ];
    vi.mocked(getThreadMessages).mockResolvedValue(ok({ conversationId: 'c1', messages }) as never);
    vi.mocked(getConsumptionHorizon).mockResolvedValue(ok({ lastReadMessageId: 'm1' }) as never);
    const res = await getThreadTool.handler(baseInput, ctx);
    expect(res.success).toBe(true);
    if (res.success) expect(res.data.unreadCount).toBe(1);
  });

  it('treats all as unread when last-read message not in window', async () => {
    const messages = [{ id: 'm2', timestamp: '2026-01-02T00:00:00Z', isFromMe: false }];
    vi.mocked(getThreadMessages).mockResolvedValue(ok({ conversationId: 'c1', messages }) as never);
    vi.mocked(getConsumptionHorizon).mockResolvedValue(ok({ lastReadMessageId: 'mX' }) as never);
    const res = await getThreadTool.handler(baseInput, ctx);
    expect(res.success).toBe(true);
    if (res.success) expect(res.data.unreadCount).toBe(1);
  });

  it('treats all as unread when there is no consumption horizon', async () => {
    const messages = [{ id: 'm1', timestamp: '2026-01-01T00:00:00Z', isFromMe: false }];
    vi.mocked(getThreadMessages).mockResolvedValue(ok({ conversationId: 'c1', messages }) as never);
    vi.mocked(getConsumptionHorizon).mockResolvedValue(ok({ lastReadMessageId: undefined }) as never);
    const res = await getThreadTool.handler(baseInput, ctx);
    expect(res.success).toBe(true);
    if (res.success) expect(res.data.unreadCount).toBe(1);
  });

  it('marks as read using newest message (desc order)', async () => {
    const messages = [
      { id: 'm2', timestamp: '2026-01-02T00:00:00Z', isFromMe: false },
      { id: 'm1', timestamp: '2026-01-01T00:00:00Z', isFromMe: false },
    ];
    vi.mocked(getThreadMessages).mockResolvedValue(ok({ conversationId: 'c1', messages }) as never);
    vi.mocked(getConsumptionHorizon).mockResolvedValue(ok({ lastReadMessageId: 'm1' }) as never);
    vi.mocked(markAsRead).mockResolvedValue(ok({}) as never);
    const res = await getThreadTool.handler({ ...baseInput, markRead: true, order: 'desc' }, ctx);
    expect(res.success).toBe(true);
    expect(vi.mocked(markAsRead)).toHaveBeenCalledWith('c1', 'm2');
    if (res.success) expect(res.data.markedAsRead).toBe(true);
  });

  it('marks as read using newest message (asc order)', async () => {
    const messages = [
      { id: 'm1', timestamp: '2026-01-01T00:00:00Z', isFromMe: false },
      { id: 'm2', timestamp: '2026-01-02T00:00:00Z', isFromMe: false },
    ];
    vi.mocked(getThreadMessages).mockResolvedValue(ok({ conversationId: 'c1', messages }) as never);
    vi.mocked(getConsumptionHorizon).mockResolvedValue(ok({ lastReadMessageId: undefined }) as never);
    vi.mocked(markAsRead).mockResolvedValue(err(createError(ErrorCode.API_ERROR, 'x')) as never);
    const res = await getThreadTool.handler({ ...baseInput, markRead: true, order: 'asc' }, ctx);
    expect(res.success).toBe(true);
    expect(vi.mocked(markAsRead)).toHaveBeenCalledWith('c1', 'm2');
    if (res.success) expect(res.data.markedAsRead).toBe(false);
  });

  it('skips marking read when there are no messages', async () => {
    vi.mocked(getThreadMessages).mockResolvedValue(ok({ conversationId: 'c1', messages: [] }) as never);
    vi.mocked(getConsumptionHorizon).mockResolvedValue(ok({ lastReadMessageId: undefined }) as never);
    const res = await getThreadTool.handler({ ...baseInput, markRead: true }, ctx);
    expect(res.success).toBe(true);
    expect(vi.mocked(markAsRead)).not.toHaveBeenCalled();
  });
});
