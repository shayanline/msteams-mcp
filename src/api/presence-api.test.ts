import { describe, it, expect, vi, beforeEach } from 'vitest';
import { ok } from '../types/result.js';

vi.mock('../utils/http.js', () => ({ httpRequest: vi.fn() }));
vi.mock('../utils/auth-guards.js', () => ({
  requireSkypeSpacesAuth: vi.fn(),
  getRegionConfig: vi.fn(),
}));

import { httpRequest } from '../utils/http.js';
import { requireSkypeSpacesAuth, getRegionConfig } from '../utils/auth-guards.js';
import { getPresence } from './presence-api.js';

const mockHttp = vi.mocked(httpRequest);
const mockAuth = vi.mocked(requireSkypeSpacesAuth);
const mockRegion = vi.mocked(getRegionConfig);
const httpOk = (data: unknown) => ok({ status: 200, headers: new Headers(), data } as never);

beforeEach(() => {
  vi.clearAllMocks();
  mockAuth.mockReturnValue(ok({ skypeToken: 'sk', spacesToken: 'sp' }) as never);
  mockRegion.mockReturnValue({ regionPartition: 'emea-02', teamsBaseUrl: 'https://teams.microsoft.com' } as never);
});

describe('getPresence', () => {
  it('normalises ids to MRIs and parses presence + out of office', async () => {
    mockHttp.mockResolvedValueOnce(httpOk([
      { mri: '8:orgid:a', presence: { availability: 'Busy', activity: 'InACall', deviceType: 'Desktop', calendarData: { isOutOfOffice: true, outOfOfficeNote: { message: '  Away  ', expiry: '2026-07-01' } } } },
    ]));
    const res = await getPresence(['8:orgid:a', 'b1c2@tenant', 'cccc']);
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value[0]).toMatchObject({ id: '8:orgid:a', availability: 'Busy', activity: 'InACall', isOutOfOffice: true, outOfOfficeMessage: 'Away', outOfOfficeExpiry: '2026-07-01' });

    const body = JSON.parse((mockHttp.mock.calls[0][1] as { body: string }).body);
    expect(body).toEqual([
      { mri: '8:orgid:a', source: 'ups' },
      { mri: '8:orgid:b1c2', source: 'ups' },
      { mri: '8:orgid:cccc', source: 'ups' },
    ]);
    expect(mockHttp.mock.calls[0][0]).toContain('/ups/emea/v1/presence/getpresence/');
  });

  it('defaults availability/activity and omits empty oof note', async () => {
    mockHttp.mockResolvedValueOnce(httpOk([{ mri: '8:orgid:a', presence: { calendarData: { outOfOfficeNote: { message: '   ' } } } }]));
    const res = await getPresence(['8:orgid:a']);
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value[0]).toMatchObject({ availability: 'Unknown', activity: 'Unknown', isOutOfOffice: false, outOfOfficeMessage: undefined });
  });

  it('fails when not authenticated', async () => {
    mockAuth.mockReturnValueOnce({ ok: false, error: { code: 'AUTH_REQUIRED' } } as never);
    const res = await getPresence(['8:orgid:a']);
    expect(res.ok).toBe(false);
    expect(mockHttp).not.toHaveBeenCalled();
  });

  it('fails when region cannot be determined', async () => {
    mockRegion.mockReturnValueOnce(null as never);
    const res = await getPresence(['8:orgid:a']);
    expect(res.ok).toBe(false);
    expect(mockHttp).not.toHaveBeenCalled();
  });

  it('propagates an http failure', async () => {
    mockHttp.mockResolvedValueOnce({ ok: false, error: { code: 'API_ERROR' } } as never);
    const res = await getPresence(['8:orgid:a']);
    expect(res.ok).toBe(false);
  });

  it('defaults when presence and calendarData are absent', async () => {
    mockHttp.mockResolvedValueOnce(httpOk([{ mri: '8:orgid:a' }]));
    const res = await getPresence(['8:orgid:a']);
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value[0]).toMatchObject({ id: '8:orgid:a', availability: 'Unknown', activity: 'Unknown', isOutOfOffice: false });
  });

  it('keeps an already-prefixed channel/bot MRI and tolerates a non-array payload', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ notAnArray: true }));
    const res = await getPresence(['28:bot-id']);
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value).toEqual([]);
    const body = JSON.parse((mockHttp.mock.calls[0][1] as { body: string }).body);
    expect(body).toEqual([{ mri: '28:bot-id', source: 'ups' }]);
  });
});
