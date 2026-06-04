import { describe, it, expect, vi, beforeEach } from 'vitest';
import { ok, err } from '../types/result.js';
import { ErrorCode, createError } from '../types/errors.js';

vi.mock('../utils/http.js', () => ({ httpRequest: vi.fn() }));
vi.mock('../utils/auth-guards.js', () => ({
  requireSkypeSpacesAuth: vi.fn(),
  getRegionConfig: vi.fn(),
}));

import { httpRequest } from '../utils/http.js';
import { requireSkypeSpacesAuth, getRegionConfig } from '../utils/auth-guards.js';
import { listTeamTags } from './tags-api.js';

const mockHttp = vi.mocked(httpRequest);
const mockAuth = vi.mocked(requireSkypeSpacesAuth);
const mockRegion = vi.mocked(getRegionConfig);
const httpOk = (data: unknown) => ok({ status: 200, headers: new Headers(), data } as never);

beforeEach(() => {
  vi.clearAllMocks();
  mockAuth.mockReturnValue(ok({ skypeToken: 'sk', spacesToken: 'sp' }) as never);
  mockRegion.mockReturnValue({ regionPartition: 'amer-02', hasPartition: true, teamsBaseUrl: 'https://teams.microsoft.com' } as never);
});

describe('listTeamTags', () => {
  it('propagates an auth failure without calling http', async () => {
    mockAuth.mockReturnValue(err(createError(ErrorCode.AUTH_REQUIRED, 'no auth')) as never);
    const res = await listTeamTags('team1');
    expect(res.ok).toBe(false);
    expect(mockHttp).not.toHaveBeenCalled();
  });

  it('fails with AUTH_REQUIRED when region config is unavailable', async () => {
    mockRegion.mockReturnValue(null as never);
    const res = await listTeamTags('team1');
    expect(res.ok).toBe(false);
    if (res.ok) return;
    expect(res.error.code).toBe(ErrorCode.AUTH_REQUIRED);
    expect(mockHttp).not.toHaveBeenCalled();
  });

  it('propagates an http failure', async () => {
    mockHttp.mockResolvedValueOnce(err(createError(ErrorCode.NETWORK_ERROR, 'boom')) as never);
    const res = await listTeamTags('team1');
    expect(res.ok).toBe(false);
  });

  it('maps tags when the API returns a top-level array', async () => {
    mockHttp.mockResolvedValueOnce(httpOk([
      { id: 't1', displayName: 'engineering', memberCount: 5, tagType: 'standard' },
    ]));
    const res = await listTeamTags('team1');
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.count).toBe(1);
    expect(res.value.tags[0]).toEqual({ id: 't1', displayName: 'engineering', memberCount: 5, tagType: 'standard' });
  });

  it('maps tags from a data.value array and applies defaults', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ value: [{}] }));
    const res = await listTeamTags('team1');
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.count).toBe(1);
    expect(res.value.tags[0]).toEqual({ id: '', displayName: '', memberCount: 0, tagType: 'standard' });
  });

  it('returns an empty list when data.value is absent', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({}));
    const res = await listTeamTags('team1');
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value).toEqual({ count: 0, tags: [] });
  });
});
