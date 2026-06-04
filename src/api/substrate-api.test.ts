import { describe, it, expect, vi, beforeEach } from 'vitest';
import { ok, err } from '../types/result.js';
import { ErrorCode, createError } from '../types/errors.js';

vi.mock('../utils/http.js', () => ({ httpRequest: vi.fn() }));
vi.mock('../utils/auth-guards.js', () => ({
  requireSubstrateTokenAsync: vi.fn(),
  getTenantId: vi.fn(),
  getTeamsBaseUrl: vi.fn(),
  handleSubstrateError: vi.fn((r: unknown) => r),
}));
vi.mock('../utils/parsers.js', () => ({
  parseSearchResults: vi.fn(),
  parseEmailSearchResults: vi.fn(),
  parsePeopleResults: vi.fn(),
  parseChannelResults: vi.fn(),
  filterChannelsByName: vi.fn(),
}));
vi.mock('./csa-api.js', () => ({ getMyTeamsAndChannels: vi.fn() }));

import { httpRequest } from '../utils/http.js';
import {
  requireSubstrateTokenAsync, getTenantId, getTeamsBaseUrl, handleSubstrateError,
} from '../utils/auth-guards.js';
import {
  parseSearchResults, parseEmailSearchResults, parsePeopleResults, parseChannelResults, filterChannelsByName,
} from '../utils/parsers.js';
import { getMyTeamsAndChannels } from './csa-api.js';
import {
  searchMessages, searchEmails, searchPeople, getFrequentContacts, searchChannels,
} from './substrate-api.js';

const mockHttp = vi.mocked(httpRequest);
const mockToken = vi.mocked(requireSubstrateTokenAsync);
const mockTenant = vi.mocked(getTenantId);
const mockBaseUrl = vi.mocked(getTeamsBaseUrl);
const mockHandle = vi.mocked(handleSubstrateError);
const mockSearch = vi.mocked(parseSearchResults);
const mockEmail = vi.mocked(parseEmailSearchResults);
const mockPeople = vi.mocked(parsePeopleResults);
const mockChannels = vi.mocked(parseChannelResults);
const mockFilter = vi.mocked(filterChannelsByName);
const mockMyTeams = vi.mocked(getMyTeamsAndChannels);
const httpOk = (data: unknown) => ok({ status: 200, headers: new Headers(), data } as never);
const lastBody = () => JSON.parse((mockHttp.mock.calls.at(-1)![1] as { body: string }).body);

beforeEach(() => {
  vi.clearAllMocks();
  mockToken.mockResolvedValue(ok('tok') as never);
  mockTenant.mockReturnValue('tid');
  mockBaseUrl.mockReturnValue('https://teams.microsoft.com');
  mockHandle.mockImplementation(((r: unknown) => r) as never);
});

describe('searchMessages', () => {
  it('propagates a token failure without calling http', async () => {
    mockToken.mockResolvedValue(err(createError(ErrorCode.AUTH_REQUIRED, 'no token')) as never);
    const res = await searchMessages('hello');
    expect(res.ok).toBe(false);
    expect(mockHttp).not.toHaveBeenCalled();
  });

  it('routes an http failure through handleSubstrateError', async () => {
    mockHttp.mockResolvedValueOnce(err(createError(ErrorCode.NETWORK_ERROR, 'boom')) as never);
    const res = await searchMessages('hi');
    expect(res.ok).toBe(false);
    expect(mockHandle).toHaveBeenCalled();
  });

  it('returns results and computes hasMore from total', async () => {
    mockSearch.mockReturnValue({ results: [{ id: 'a' }, { id: 'b' }, { id: 'c' }], total: 10 } as never);
    mockHttp.mockResolvedValueOnce(httpOk({ EntitySets: [] }));
    const res = await searchMessages('hi');
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.results).toHaveLength(3);
    expect(res.value.pagination).toMatchObject({ from: 0, size: 25, returned: 3, total: 10, hasMore: true });
    expect(lastBody().entityRequests[0].query.queryString).toContain('hi AND NOT');
  });

  it('limits results via maxResults and computes hasMore without total', async () => {
    mockSearch.mockReturnValue({ results: [{ id: 'a' }, { id: 'b' }, { id: 'c' }, { id: 'd' }, { id: 'e' }], total: undefined } as never);
    mockHttp.mockResolvedValueOnce(httpOk({ EntitySets: [] }));
    const res = await searchMessages('hi', { from: 0, size: 5, maxResults: 2 });
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.results).toHaveLength(2);
    expect(res.value.pagination.hasMore).toBe(true);
    expect(res.value.pagination.total).toBeUndefined();
  });
});

describe('searchEmails', () => {
  it('propagates a token failure', async () => {
    mockToken.mockResolvedValue(err(createError(ErrorCode.AUTH_REQUIRED, 'x')) as never);
    const res = await searchEmails('q');
    expect(res.ok).toBe(false);
    expect(mockHttp).not.toHaveBeenCalled();
  });

  it('routes an http failure through handleSubstrateError', async () => {
    mockHttp.mockResolvedValueOnce(err(createError(ErrorCode.NETWORK_ERROR, 'boom')) as never);
    const res = await searchEmails('q');
    expect(res.ok).toBe(false);
    expect(mockHandle).toHaveBeenCalled();
  });

  it('parses results with EntitySets array and hasMore from total', async () => {
    mockEmail.mockReturnValue({ results: [{ id: 'm1' }], total: 50, filteredCount: 2 } as never);
    mockHttp.mockResolvedValueOnce(httpOk({ EntitySets: [{}] }));
    const res = await searchEmails('q', { from: 0, size: 25 });
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.filteredCount).toBe(2);
    expect(res.value.pagination).toMatchObject({ total: 50, hasMore: true });
  });

  it('handles non-array EntitySets and hasMore without total', async () => {
    mockEmail.mockReturnValue({ results: [{ id: 'm1' }, { id: 'm2' }], total: undefined, filteredCount: undefined } as never);
    mockHttp.mockResolvedValueOnce(httpOk({ EntitySets: 'nope' }));
    const res = await searchEmails('q', { size: 2, maxResults: 1 });
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.results).toHaveLength(1);
    expect(res.value.pagination.hasMore).toBe(true);
  });
});

describe('searchPeople', () => {
  it('propagates a token failure', async () => {
    mockToken.mockResolvedValue(err(createError(ErrorCode.AUTH_REQUIRED, 'x')) as never);
    const res = await searchPeople('bob');
    expect(res.ok).toBe(false);
    expect(mockHttp).not.toHaveBeenCalled();
  });

  it('routes an http failure through handleSubstrateError', async () => {
    mockHttp.mockResolvedValueOnce(err(createError(ErrorCode.NETWORK_ERROR, 'boom')) as never);
    const res = await searchPeople('bob');
    expect(res.ok).toBe(false);
    expect(mockHandle).toHaveBeenCalled();
  });

  it('returns parsed people and the count', async () => {
    mockPeople.mockReturnValue([{ mri: '1' }, { mri: '2' }] as never);
    mockHttp.mockResolvedValueOnce(httpOk({ Groups: [] }));
    const res = await searchPeople('bob', 5);
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.returned).toBe(2);
    expect(lastBody().EntityRequests[0].Size).toBe(5);
  });
});

describe('getFrequentContacts', () => {
  it('propagates a token failure', async () => {
    mockToken.mockResolvedValue(err(createError(ErrorCode.AUTH_REQUIRED, 'x')) as never);
    const res = await getFrequentContacts();
    expect(res.ok).toBe(false);
    expect(mockHttp).not.toHaveBeenCalled();
  });

  it('routes an http failure through handleSubstrateError', async () => {
    mockHttp.mockResolvedValueOnce(err(createError(ErrorCode.NETWORK_ERROR, 'boom')) as never);
    const res = await getFrequentContacts();
    expect(res.ok).toBe(false);
    expect(mockHandle).toHaveBeenCalled();
  });

  it('returns parsed contacts with the default limit', async () => {
    mockPeople.mockReturnValue([{ mri: '1' }] as never);
    mockHttp.mockResolvedValueOnce(httpOk({ Groups: [] }));
    const res = await getFrequentContacts();
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.returned).toBe(1);
    expect(lastBody().EntityRequests[0].Size).toBe(50);
  });
});

describe('searchChannels', () => {
  it('propagates a token failure', async () => {
    mockToken.mockResolvedValue(err(createError(ErrorCode.AUTH_REQUIRED, 'x')) as never);
    const res = await searchChannels('eng');
    expect(res.ok).toBe(false);
    expect(mockHttp).not.toHaveBeenCalled();
  });

  it('merges user channels and org-wide results, deduping and marking membership', async () => {
    // org-wide http call returns channels c1 (member) and c3 (not member)
    mockChannels.mockReturnValue([
      { channelId: 'c1', isMember: false },
      { channelId: 'c3', isMember: false },
    ] as never);
    mockHttp.mockResolvedValueOnce(httpOk({ Groups: [] }));
    // user's teams contain c1 and c2
    mockMyTeams.mockResolvedValue(ok({ teams: [
      { channels: [{ channelId: 'c1' }, { channelId: 'c2' }] },
    ] }) as never);
    // filterChannelsByName returns the matching user channel c1
    mockFilter.mockReturnValue([{ channelId: 'c1' }] as never);

    const res = await searchChannels('eng', 10);
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    const ids = res.value.results.map(c => c.channelId);
    expect(ids).toEqual(['c1', 'c3']);
    // c1 is kept from the user's matching channels (dedup drops the org copy)
    const c1 = res.value.results.find(c => c.channelId === 'c1')!;
    expect(c1.isMember).toBeUndefined();
    // c3 only comes from org-wide results and is marked as not a member
    const c3 = res.value.results.find(c => c.channelId === 'c3')!;
    expect(c3.isMember).toBe(false);
    expect(res.value.returned).toBe(2);
  });

  it('handles a failed teams lookup and a failed org-wide search', async () => {
    // org-wide http fails -> searchChannelsOrgWide returns failure via handleSubstrateError
    mockHttp.mockResolvedValueOnce(err(createError(ErrorCode.NETWORK_ERROR, 'boom')) as never);
    mockMyTeams.mockResolvedValue(err(createError(ErrorCode.API_ERROR, 'nope')) as never);
    const res = await searchChannels('eng');
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.results).toEqual([]);
    expect(res.value.returned).toBe(0);
  });

  it('applies the limit to merged results', async () => {
    mockChannels.mockReturnValue([
      { channelId: 'o1', isMember: false },
      { channelId: 'o2', isMember: false },
    ] as never);
    mockHttp.mockResolvedValueOnce(httpOk({ Groups: [] }));
    mockMyTeams.mockResolvedValue(ok({ teams: [{ channels: [] }] }) as never);
    mockFilter.mockReturnValue([{ channelId: 'u1' }] as never);
    const res = await searchChannels('x', 2);
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.results.map(c => c.channelId)).toEqual(['u1', 'o1']);
  });
});
