/**
 * Unit tests for auth-guards utilities.
 *
 * The token-extractor and token-refresh dependencies are mocked so the guard
 * logic, region/tenant caching and Result wiring can be tested in isolation.
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';

vi.mock('../auth/token-extractor.js', () => ({
  getValidSubstrateToken: vi.fn(),
  extractMessageAuth: vi.fn(),
  extractCsaToken: vi.fn(),
  extractSubstrateToken: vi.fn(),
  extractSkypeSpacesToken: vi.fn(),
  extractGraphToken: vi.fn(),
  extractRegionConfig: vi.fn(),
  getUserProfile: vi.fn(),
  clearTokenCache: vi.fn(),
}));

vi.mock('../auth/token-refresh.js', () => ({
  refreshTokensViaBrowser: vi.fn(),
}));

import {
  handleSubstrateError,
  requireSubstrateTokenAsync,
  requireMessageAuth,
  requireCsaAuth,
  requireSkypeSpacesAuth,
  requireGraphAuth,
  requireMessageAuthWithConfig,
  getApiConfig,
  getRegion,
  getTeamsBaseUrl,
  getRegionConfig,
  getTenantId,
  clearRegionCache,
} from './auth-guards.js';
import { ErrorCode, createError } from '../types/errors.js';
import { type Result, ok, err } from '../types/result.js';
import {
  getValidSubstrateToken,
  extractMessageAuth,
  extractCsaToken,
  extractSubstrateToken,
  extractSkypeSpacesToken,
  extractGraphToken,
  extractRegionConfig,
  getUserProfile,
  clearTokenCache,
} from '../auth/token-extractor.js';
import { refreshTokensViaBrowser } from '../auth/token-refresh.js';

const messageAuth = { skypeToken: 'sk', authToken: 'at', userMri: '8:orgid:me' };
const regionConfig = {
  region: 'emea',
  partition: '03',
  regionPartition: 'emea-03',
  hasPartition: true,
  middleTierUrl: 'https://teams.microsoft.us/api/mt/part/emea-03',
  chatServiceUrl: 'https://teams.microsoft.us/api/chatsvc/emea',
  csaServiceUrl: 'https://teams.microsoft.us/api/csa/emea',
  teamsBaseUrl: 'https://teams.microsoft.us',
};
const futureToken = () => ({ token: 'sub', expiry: new Date(Date.now() + 60 * 60 * 1000) });
const soonToken = () => ({ token: 'sub', expiry: new Date(Date.now() + 60 * 1000) });

beforeEach(() => {
  vi.clearAllMocks();
  clearRegionCache();
});

// ============================================================================
// handleSubstrateError (existing behaviour, preserved)
// ============================================================================

describe('handleSubstrateError', () => {
  it('returns the same error result unchanged', () => {
    const result: Result<string> = err(createError(ErrorCode.API_ERROR, 'Server error'));
    const handled = handleSubstrateError(result);
    expect(handled.ok).toBe(false);
    if (!handled.ok) {
      expect(handled.error.code).toBe(ErrorCode.API_ERROR);
      expect(handled.error.message).toBe('Server error');
    }
  });

  it('returns the error for AUTH_EXPIRED and clears the token cache', () => {
    const result: Result<string> = err(createError(ErrorCode.AUTH_EXPIRED, 'Token expired'));
    const handled = handleSubstrateError(result);
    expect(handled.ok).toBe(false);
    if (!handled.ok) expect(handled.error.code).toBe(ErrorCode.AUTH_EXPIRED);
    expect(clearTokenCache).toHaveBeenCalledOnce();
  });

  it('passes through non-AUTH_EXPIRED errors without clearing cache', () => {
    const result: Result<string> = err(createError(ErrorCode.RATE_LIMITED, 'Too many requests'));
    const handled = handleSubstrateError(result);
    expect(handled.ok).toBe(false);
    if (!handled.ok) expect(handled.error.code).toBe(ErrorCode.RATE_LIMITED);
    expect(clearTokenCache).not.toHaveBeenCalled();
  });

  it('does not break if accidentally called on an ok result (defensive)', () => {
    const result: Result<string> = ok('success');
    const handled = handleSubstrateError(result);
    expect(handled.ok).toBe(true);
    if (handled.ok) expect(handled.value).toBe('success');
    expect(clearTokenCache).not.toHaveBeenCalled();
  });
});

// ============================================================================
// requireSubstrateTokenAsync
// ============================================================================

describe('requireSubstrateTokenAsync', () => {
  it('returns the existing token when no refresh is needed', async () => {
    vi.mocked(extractSubstrateToken).mockReturnValue(futureToken());
    vi.mocked(getValidSubstrateToken).mockReturnValue('valid-token');

    const result = await requireSubstrateTokenAsync();

    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toBe('valid-token');
    expect(refreshTokensViaBrowser).not.toHaveBeenCalled();
  });

  it('refreshes when no token is present and returns the new token', async () => {
    vi.mocked(extractSubstrateToken).mockReturnValue(null); // shouldRefresh -> true
    vi.mocked(refreshTokensViaBrowser).mockResolvedValue(
      ok({ method: 'http' }) as never,
    );
    vi.mocked(getValidSubstrateToken).mockReturnValue('refreshed-token');

    const result = await requireSubstrateTokenAsync();

    expect(refreshTokensViaBrowser).toHaveBeenCalled();
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toBe('refreshed-token');
  });

  it('refreshes when token is close to expiry', async () => {
    vi.mocked(extractSubstrateToken).mockReturnValue(soonToken());
    vi.mocked(refreshTokensViaBrowser).mockResolvedValue(ok({ method: 'http' }) as never);
    vi.mocked(getValidSubstrateToken).mockReturnValue('refreshed-token');

    const result = await requireSubstrateTokenAsync();

    expect(refreshTokensViaBrowser).toHaveBeenCalled();
    expect(result.ok).toBe(true);
  });

  it('continues with an existing token when refresh fails', async () => {
    vi.mocked(extractSubstrateToken).mockReturnValue(null);
    vi.mocked(refreshTokensViaBrowser).mockResolvedValue(
      err(createError(ErrorCode.NETWORK_ERROR, 'boom')),
    );
    vi.mocked(getValidSubstrateToken).mockReturnValue('still-valid');

    const result = await requireSubstrateTokenAsync();

    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toBe('still-valid');
  });

  it('returns AUTH_EXPIRED when refresh succeeds but no token can be read', async () => {
    vi.mocked(extractSubstrateToken).mockReturnValue(null);
    vi.mocked(refreshTokensViaBrowser).mockResolvedValue(ok({ method: 'http' }) as never);
    vi.mocked(getValidSubstrateToken).mockReturnValue(null);

    const result = await requireSubstrateTokenAsync();

    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.code).toBe(ErrorCode.AUTH_EXPIRED);
  });

  it('returns AUTH_EXPIRED when refresh fails and no token is available', async () => {
    vi.mocked(extractSubstrateToken).mockReturnValue(null);
    vi.mocked(refreshTokensViaBrowser).mockResolvedValue(
      err(createError(ErrorCode.AUTH_EXPIRED, 'expired')),
    );
    vi.mocked(getValidSubstrateToken).mockReturnValue(null);

    const result = await requireSubstrateTokenAsync();

    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.code).toBe(ErrorCode.AUTH_EXPIRED);
  });
});

// ============================================================================
// requireMessageAuth / requireCsaAuth / requireSkypeSpacesAuth / requireGraphAuth
// ============================================================================

describe('requireMessageAuth', () => {
  it('returns auth info when present', () => {
    vi.mocked(extractMessageAuth).mockReturnValue(messageAuth);
    const result = requireMessageAuth();
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toEqual(messageAuth);
  });

  it('returns AUTH_REQUIRED when absent', () => {
    vi.mocked(extractMessageAuth).mockReturnValue(null);
    const result = requireMessageAuth();
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.code).toBe(ErrorCode.AUTH_REQUIRED);
  });
});

describe('requireCsaAuth', () => {
  it('returns combined auth + csa token when both present', () => {
    vi.mocked(extractMessageAuth).mockReturnValue(messageAuth);
    vi.mocked(extractCsaToken).mockReturnValue('csa');
    const result = requireCsaAuth();
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value.csaToken).toBe('csa');
      expect(result.value.auth).toEqual(messageAuth);
    }
  });

  it('returns AUTH_REQUIRED when csa token missing', () => {
    vi.mocked(extractMessageAuth).mockReturnValue(messageAuth);
    vi.mocked(extractCsaToken).mockReturnValue(null);
    const result = requireCsaAuth();
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.code).toBe(ErrorCode.AUTH_REQUIRED);
  });

  it('returns AUTH_REQUIRED when message auth missing', () => {
    vi.mocked(extractMessageAuth).mockReturnValue(null);
    vi.mocked(extractCsaToken).mockReturnValue('csa');
    const result = requireCsaAuth();
    expect(result.ok).toBe(false);
  });
});

describe('requireSkypeSpacesAuth', () => {
  it('returns skype + spaces tokens when present', () => {
    vi.mocked(extractMessageAuth).mockReturnValue(messageAuth);
    vi.mocked(extractSkypeSpacesToken).mockReturnValue('spaces');
    const result = requireSkypeSpacesAuth();
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value.skypeToken).toBe('sk');
      expect(result.value.spacesToken).toBe('spaces');
    }
  });

  it('returns AUTH_REQUIRED when the spaces token is missing', () => {
    vi.mocked(extractMessageAuth).mockReturnValue(messageAuth);
    vi.mocked(extractSkypeSpacesToken).mockReturnValue(null);
    const result = requireSkypeSpacesAuth();
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.code).toBe(ErrorCode.AUTH_REQUIRED);
  });
});

describe('requireGraphAuth', () => {
  it('returns the graph token when present', () => {
    vi.mocked(extractGraphToken).mockReturnValue('graph');
    const result = requireGraphAuth();
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toBe('graph');
  });

  it('returns AUTH_REQUIRED when the graph token is missing', () => {
    vi.mocked(extractGraphToken).mockReturnValue(null);
    const result = requireGraphAuth();
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.code).toBe(ErrorCode.AUTH_REQUIRED);
  });
});

// ============================================================================
// Region / API config (with caching)
// ============================================================================

describe('region and API config', () => {
  it('returns region and base URL from session config', () => {
    vi.mocked(extractRegionConfig).mockReturnValue(regionConfig);
    expect(getRegion()).toBe('emea');
    expect(getTeamsBaseUrl()).toBe('https://teams.microsoft.us');
    // Cached: extractRegionConfig only called once across both reads.
    expect(extractRegionConfig).toHaveBeenCalledTimes(1);
  });

  it('falls back to defaults when no config is available', () => {
    vi.mocked(extractRegionConfig).mockReturnValue(null);
    expect(getRegion()).toBe('amer');
    expect(getTeamsBaseUrl()).toBe('https://teams.microsoft.com');
  });

  it('getRegionConfig returns the full config or null', () => {
    vi.mocked(extractRegionConfig).mockReturnValue(regionConfig);
    expect(getRegionConfig()).toEqual(regionConfig);

    clearRegionCache();
    vi.mocked(extractRegionConfig).mockReturnValue(null);
    expect(getRegionConfig()).toBeNull();
  });

  it('getApiConfig combines region and base URL', () => {
    vi.mocked(extractRegionConfig).mockReturnValue(regionConfig);
    expect(getApiConfig()).toEqual({ region: 'emea', baseUrl: 'https://teams.microsoft.us' });
  });
});

describe('requireMessageAuthWithConfig', () => {
  it('returns auth plus region config on success', () => {
    vi.mocked(extractMessageAuth).mockReturnValue(messageAuth);
    vi.mocked(extractRegionConfig).mockReturnValue(regionConfig);
    const result = requireMessageAuthWithConfig();
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value.auth).toEqual(messageAuth);
      expect(result.value.region).toBe('emea');
      expect(result.value.baseUrl).toBe('https://teams.microsoft.us');
    }
  });

  it('propagates the auth error when message auth is missing', () => {
    vi.mocked(extractMessageAuth).mockReturnValue(null);
    const result = requireMessageAuthWithConfig();
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.code).toBe(ErrorCode.AUTH_REQUIRED);
  });
});

// ============================================================================
// getTenantId (with caching)
// ============================================================================

describe('getTenantId', () => {
  it('returns the tenant id from the user profile and caches it', () => {
    vi.mocked(getUserProfile).mockReturnValue({ tenantId: 'tenant-123' } as never);
    expect(getTenantId()).toBe('tenant-123');
    // Second call hits the cache.
    expect(getTenantId()).toBe('tenant-123');
    expect(getUserProfile).toHaveBeenCalledTimes(1);
  });

  it('returns null when no profile is available', () => {
    vi.mocked(getUserProfile).mockReturnValue(null);
    expect(getTenantId()).toBeNull();
  });
});
