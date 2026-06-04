import { describe, it, expect, vi, beforeEach } from 'vitest';
import { ok } from '../types/result.js';

vi.mock('../utils/http.js', () => ({ httpRequest: vi.fn() }));
vi.mock('../utils/auth-guards.js', () => ({
  requireMessageAuth: vi.fn(),
  requireSubstrateTokenAsync: vi.fn(),
  getTenantId: vi.fn(),
  handleSubstrateError: vi.fn((r) => r),
}));

import { httpRequest } from '../utils/http.js';
import { requireMessageAuth, requireSubstrateTokenAsync, getTenantId } from '../utils/auth-guards.js';
import { getSharedFiles } from './files-api.js';

const mockHttp = vi.mocked(httpRequest);
const mockMsgAuth = vi.mocked(requireMessageAuth);
const mockSub = vi.mocked(requireSubstrateTokenAsync);
const mockTenant = vi.mocked(getTenantId);
const httpOk = (data: unknown) => ok({ status: 200, headers: new Headers(), data } as never);

beforeEach(() => {
  vi.clearAllMocks();
  mockMsgAuth.mockReturnValue(ok({ skypeToken: 'sk', authToken: 'at', userMri: '8:orgid:dde37b63-a4ac-4edb-a7b0-385263022300' }) as never);
  mockSub.mockResolvedValue(ok('sub-token') as never);
  mockTenant.mockReturnValue('tenant-id');
});

describe('getSharedFiles', () => {
  it('maps file and link items', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({
      Items: [
        { ItemType: 'File', FileData: { FileName: 'a.pdf', FileExtension: 'pdf', FileUrl: 'http://sp/a.pdf', SizeInBytes: 100, PreviewUrl: 'http://p' }, SharedByDisplayName: 'Alice', SharedBySmtp: 'a@x.com', SharedDateTime: '2026-01-01' },
        { ItemType: 'Link', WeblinkData: { WebUrl: 'http://link', Title: 'Doc' }, SharedByDisplayName: 'Bob' },
      ],
      SkipToken: 'next',
    }));
    const res = await getSharedFiles('19:c@thread.tacv2');
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.returned).toBe(2);
    expect(res.value.skipToken).toBe('next');
    expect(res.value.files[0]).toMatchObject({ itemType: 'File', fileName: 'a.pdf', webUrl: 'http://sp/a.pdf', sharedBy: 'Alice' });
    expect(res.value.files[1]).toMatchObject({ itemType: 'Link', webUrl: 'http://link', title: 'Doc' });
    const url = mockHttp.mock.calls[0][0] as string;
    expect(url).toContain('/AllFiles/api/users(');
  });

  it('returns empty when there are no items', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ Items: [] }));
    const res = await getSharedFiles('19:c@thread.tacv2');
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.files).toEqual([]);
    expect(res.value.returned).toBe(0);
  });

  it('fails without message auth', async () => {
    mockMsgAuth.mockReturnValueOnce({ ok: false, error: { code: 'AUTH_REQUIRED' } } as never);
    const res = await getSharedFiles('19:c@thread.tacv2');
    expect(res.ok).toBe(false);
  });

  it('fails when the substrate token is unavailable', async () => {
    mockSub.mockResolvedValueOnce({ ok: false, error: { code: 'AUTH_REQUIRED' } } as never);
    const res = await getSharedFiles('19:c@thread.tacv2');
    expect(res.ok).toBe(false);
  });
});
