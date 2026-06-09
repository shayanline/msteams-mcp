import { describe, it, expect, vi, beforeEach } from 'vitest';
import { ok, err } from '../types/result.js';
import { ErrorCode, createError } from '../types/errors.js';

vi.mock('../utils/http.js', () => ({ httpRequest: vi.fn() }));
vi.mock('../utils/auth-guards.js', () => ({
  requireSkypeSpacesAuth: vi.fn(),
  getRegionConfig: vi.fn(),
}));
vi.mock('./csa-api.js', () => ({ getMyTeamsAndChannels: vi.fn() }));

import { httpRequest } from '../utils/http.js';
import { requireSkypeSpacesAuth, getRegionConfig } from '../utils/auth-guards.js';
import { getMyTeamsAndChannels } from './csa-api.js';
import { createChannel, deleteChannel } from './channels-api.js';

const mockHttp = vi.mocked(httpRequest);
const mockAuth = vi.mocked(requireSkypeSpacesAuth);
const mockRegion = vi.mocked(getRegionConfig);
const mockTeams = vi.mocked(getMyTeamsAndChannels);
const httpOk = (data: unknown) => ok({ status: 200, headers: new Headers(), data } as never);

const TEAM = {
  teamId: '19:root@thread.tacv2',          // parseTeamsList sets teamId to the thread id
  teamName: 'CRM',
  threadId: '19:root@thread.tacv2',
  channels: [{ channelId: '19:gen@thread.tacv2', channelName: 'General', teamName: 'CRM', teamId: 'GROUP-GUID', channelType: 'Standard', isMember: true }],
};

beforeEach(() => {
  vi.clearAllMocks();
  mockAuth.mockReturnValue(ok({ skypeToken: 'sk', spacesToken: 'sp' }) as never);
  mockRegion.mockReturnValue({ regionPartition: 'emea-02', hasPartition: true, teamsBaseUrl: 'https://teams.microsoft.com' } as never);
  mockTeams.mockResolvedValue(ok({ teams: [TEAM] }) as never);
});

describe('createChannel', () => {
  it('propagates an auth failure without calling http', async () => {
    mockAuth.mockReturnValue(err(createError(ErrorCode.AUTH_REQUIRED, 'no auth')) as never);
    const res = await createChannel('GROUP-GUID', 'New');
    expect(res.ok).toBe(false);
    expect(mockHttp).not.toHaveBeenCalled();
  });

  it('fails when the team cannot be found among the user\'s teams', async () => {
    mockTeams.mockResolvedValue(ok({ teams: [] }) as never);
    const res = await createChannel('GROUP-GUID', 'New');
    expect(res.ok).toBe(false);
    if (res.ok) return;
    expect(res.error.code).toBe(ErrorCode.NOT_FOUND);
    expect(mockHttp).not.toHaveBeenCalled();
  });

  it('resolves the team root thread id from the group GUID and posts the create request', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ value: { objectId: '19:new@thread.tacv2', displayName: 'New', description: 'd' } }));
    const res = await createChannel('GROUP-GUID', 'New', 'd');
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value).toMatchObject({ conversationId: '19:new@thread.tacv2', displayName: 'New', teamId: 'GROUP-GUID', membershipType: 'standard' });

    const [url, options] = mockHttp.mock.calls[0] as [string, { method: string; body: string }];
    // URL keys on the team root thread id, not the group GUID.
    expect(url).toContain('/beta/teams/19%3Aroot%40thread.tacv2/channels');
    expect(options.method).toBe('POST');
    expect(JSON.parse(options.body)).toEqual({
      displayName: 'New',
      description: 'd',
      groupId: 'GROUP-GUID',
      channelType: 'Standard',
      chatModalityType: 'Conversational',
    });
  });

  it('sends channelType Private for a private channel', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ value: { objectId: '19:priv@thread.tacv2' } }));
    const res = await createChannel('GROUP-GUID', 'Secret', '', 'private');
    expect(res.ok).toBe(true);
    const body = JSON.parse((mockHttp.mock.calls[0][1] as { body: string }).body);
    expect(body.channelType).toBe('Private');
  });

  it('errors when the response has no objectId', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ value: {} }));
    const res = await createChannel('GROUP-GUID', 'New');
    expect(res.ok).toBe(false);
  });

  it('propagates an http failure', async () => {
    mockHttp.mockResolvedValueOnce(err(createError(ErrorCode.NETWORK_ERROR, 'boom')) as never);
    const res = await createChannel('GROUP-GUID', 'New');
    expect(res.ok).toBe(false);
  });
});

describe('deleteChannel', () => {
  it('propagates an auth failure without calling http', async () => {
    mockAuth.mockReturnValue(err(createError(ErrorCode.AUTH_REQUIRED, 'no auth')) as never);
    const res = await deleteChannel('GROUP-GUID', '19:chan@thread.tacv2');
    expect(res.ok).toBe(false);
    expect(mockHttp).not.toHaveBeenCalled();
  });

  it('fails when the team cannot be found', async () => {
    mockTeams.mockResolvedValue(ok({ teams: [] }) as never);
    const res = await deleteChannel('GROUP-GUID', '19:chan@thread.tacv2');
    expect(res.ok).toBe(false);
    if (res.ok) return;
    expect(res.error.code).toBe(ErrorCode.NOT_FOUND);
    expect(mockHttp).not.toHaveBeenCalled();
  });

  it('sends a DELETE to the channel URL keyed on the team root thread id', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({}));
    const res = await deleteChannel('GROUP-GUID', '19:chan@thread.tacv2');
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value).toEqual({ conversationId: '19:chan@thread.tacv2', teamId: 'GROUP-GUID' });

    const [url, options] = mockHttp.mock.calls[0] as [string, { method: string }];
    expect(options.method).toBe('DELETE');
    expect(url).toContain('/beta/teams/19%3Aroot%40thread.tacv2/channels/19%3Achan%40thread.tacv2');
  });

  it('propagates an http failure', async () => {
    mockHttp.mockResolvedValueOnce(err(createError(ErrorCode.NETWORK_ERROR, 'boom')) as never);
    const res = await deleteChannel('GROUP-GUID', '19:chan@thread.tacv2');
    expect(res.ok).toBe(false);
  });
});
