/**
 * Tests for the token refresh orchestrator (HTTP-first, browser fallback).
 *
 * The HTTP refresh, token extraction, logger and the dynamically imported
 * browser modules are all mocked so the orchestration logic can be exercised
 * without a real browser or network.
 */

import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { ErrorCode } from './../types/errors.js';

vi.mock('./token-extractor.js', () => ({
  extractSubstrateToken: vi.fn(),
  clearTokenCache: vi.fn(),
}));

vi.mock('./token-refresh-http.js', () => ({
  refreshTokensViaHttp: vi.fn(),
}));

vi.mock('../browser/context.js', () => ({
  createBrowserContext: vi.fn(),
  closeBrowser: vi.fn(),
}));

vi.mock('../browser/auth.js', () => ({
  ensureAuthenticated: vi.fn(),
}));

vi.mock('../utils/logger.js', () => ({
  warn: vi.fn(),
  error: vi.fn(),
  info: vi.fn(),
  debug: vi.fn(),
}));

import { refreshTokensViaBrowser } from './token-refresh.js';
import { extractSubstrateToken, clearTokenCache } from './token-extractor.js';
import { refreshTokensViaHttp } from './token-refresh-http.js';
import { createBrowserContext, closeBrowser } from '../browser/context.js';
import { ensureAuthenticated } from '../browser/auth.js';
import { ok, err } from './../types/result.js';
import { createError } from './../types/errors.js';

const tokenInfo = (msAhead: number) => ({
  token: 'tok',
  expiry: new Date(Date.now() + msAhead),
});

const httpOk = () =>
  ok({ tokensRefreshed: 4, skypeTokenRefreshed: true, refreshTokenRotated: true });

const fakeManager = () => ({ page: {}, context: {} }) as never;

beforeEach(() => {
  vi.clearAllMocks();
});

afterEach(() => {
  vi.restoreAllMocks();
});

describe('refreshTokensViaBrowser - guard & HTTP path', () => {
  it('returns AUTH_REQUIRED when no current token exists', async () => {
    vi.mocked(extractSubstrateToken).mockReturnValue(null);

    const result = await refreshTokensViaBrowser();

    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.code).toBe(ErrorCode.AUTH_REQUIRED);
  });

  it('succeeds via HTTP refresh and reports refreshNeeded when close to expiry', async () => {
    // before token close to expiry (< 10 min threshold)
    vi.mocked(extractSubstrateToken)
      .mockReturnValueOnce(tokenInfo(5 * 60 * 1000))
      .mockReturnValueOnce(tokenInfo(60 * 60 * 1000));
    vi.mocked(refreshTokensViaHttp).mockResolvedValue(httpOk());

    const result = await refreshTokensViaBrowser();

    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value.method).toBe('http');
      expect(result.value.refreshNeeded).toBe(true);
      expect(result.value.minutesGained).toBeGreaterThan(0);
    }
  });

  it('succeeds via HTTP refresh with refreshNeeded false when token is far from expiry', async () => {
    vi.mocked(extractSubstrateToken)
      .mockReturnValueOnce(tokenInfo(60 * 60 * 1000))
      .mockReturnValueOnce(tokenInfo(120 * 60 * 1000));
    vi.mocked(refreshTokensViaHttp).mockResolvedValue(
      ok({ tokensRefreshed: 1, skypeTokenRefreshed: false, refreshTokenRotated: false }),
    );

    const result = await refreshTokensViaBrowser();

    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value.refreshNeeded).toBe(false);
  });

  it('rejects concurrent refresh attempts with a retryable error', async () => {
    let release!: () => void;
    const gate = new Promise<void>((resolve) => { release = resolve; });

    vi.mocked(extractSubstrateToken).mockReturnValue(tokenInfo(60 * 60 * 1000));
    vi.mocked(refreshTokensViaHttp).mockImplementation(async () => {
      await gate;
      return httpOk();
    });

    const first = refreshTokensViaBrowser();
    // While the first is awaiting the gate, the flag is set.
    const second = await refreshTokensViaBrowser();

    expect(second.ok).toBe(false);
    if (!second.ok) {
      expect(second.error.code).toBe(ErrorCode.UNKNOWN);
      expect(second.error.retryable).toBe(true);
    }

    release();
    await first;
  });
});

describe('refreshTokensViaBrowser - browser fallback', () => {
  it('falls back to the browser when HTTP reports success but no valid token', async () => {
    // before, http-after invalid (null), browser-after valid
    vi.mocked(extractSubstrateToken)
      .mockReturnValueOnce(tokenInfo(5 * 60 * 1000))
      .mockReturnValueOnce(null)
      .mockReturnValueOnce(tokenInfo(60 * 60 * 1000));
    vi.mocked(refreshTokensViaHttp).mockResolvedValue(httpOk());
    vi.mocked(createBrowserContext).mockResolvedValue(fakeManager());
    // Invoke the supplied logging callback so its body is exercised.
    vi.mocked(ensureAuthenticated).mockImplementation(
      (async (_page: unknown, _ctx: unknown, logCb: (msg: string) => void) => {
        logCb('headless refresh in progress');
      }) as never,
    );
    vi.mocked(closeBrowser).mockResolvedValue(undefined as never);

    const result = await refreshTokensViaBrowser();

    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value.method).toBe('browser');
    expect(clearTokenCache).toHaveBeenCalled();
  });

  it('falls back to the browser when HTTP fails with AUTH_EXPIRED', async () => {
    vi.mocked(extractSubstrateToken)
      .mockReturnValueOnce(tokenInfo(60 * 60 * 1000))
      .mockReturnValueOnce(tokenInfo(120 * 60 * 1000));
    vi.mocked(refreshTokensViaHttp).mockResolvedValue(
      err(createError(ErrorCode.AUTH_EXPIRED, 'refresh token expired')),
    );
    vi.mocked(createBrowserContext).mockResolvedValue(fakeManager());
    vi.mocked(ensureAuthenticated).mockResolvedValue(undefined as never);
    vi.mocked(closeBrowser).mockResolvedValue(undefined as never);

    const result = await refreshTokensViaBrowser();

    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value.method).toBe('browser');
  });

  it('falls back to the browser when HTTP fails with a non-auth error', async () => {
    vi.mocked(extractSubstrateToken)
      .mockReturnValueOnce(tokenInfo(60 * 60 * 1000))
      .mockReturnValueOnce(tokenInfo(120 * 60 * 1000));
    vi.mocked(refreshTokensViaHttp).mockResolvedValue(
      err(createError(ErrorCode.NETWORK_ERROR, 'boom')),
    );
    vi.mocked(createBrowserContext).mockResolvedValue(fakeManager());
    vi.mocked(ensureAuthenticated).mockResolvedValue(undefined as never);
    vi.mocked(closeBrowser).mockResolvedValue(undefined as never);

    const result = await refreshTokensViaBrowser();

    expect(result.ok).toBe(true);
  });

  it('returns AUTH_EXPIRED when the browser yields no token', async () => {
    vi.mocked(extractSubstrateToken)
      .mockReturnValueOnce(tokenInfo(60 * 60 * 1000))
      .mockReturnValueOnce(null);
    vi.mocked(refreshTokensViaHttp).mockResolvedValue(err(createError(ErrorCode.NETWORK_ERROR, 'x')));
    vi.mocked(createBrowserContext).mockResolvedValue(fakeManager());
    vi.mocked(ensureAuthenticated).mockResolvedValue(undefined as never);
    vi.mocked(closeBrowser).mockResolvedValue(undefined as never);

    const result = await refreshTokensViaBrowser();

    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.code).toBe(ErrorCode.AUTH_EXPIRED);
  });

  it('returns AUTH_EXPIRED when close to expiry but the token was not refreshed', async () => {
    const before = tokenInfo(5 * 60 * 1000); // close to expiry
    vi.mocked(extractSubstrateToken)
      .mockReturnValueOnce(before)
      .mockReturnValueOnce({ token: 'tok', expiry: before.expiry }); // not advanced
    vi.mocked(refreshTokensViaHttp).mockResolvedValue(err(createError(ErrorCode.NETWORK_ERROR, 'x')));
    vi.mocked(createBrowserContext).mockResolvedValue(fakeManager());
    vi.mocked(ensureAuthenticated).mockResolvedValue(undefined as never);
    vi.mocked(closeBrowser).mockResolvedValue(undefined as never);

    const result = await refreshTokensViaBrowser();

    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.code).toBe(ErrorCode.AUTH_EXPIRED);
  });

  it('returns UNKNOWN and cleans up when the browser flow throws', async () => {
    vi.mocked(extractSubstrateToken).mockReturnValueOnce(tokenInfo(60 * 60 * 1000));
    vi.mocked(refreshTokensViaHttp).mockResolvedValue(err(createError(ErrorCode.NETWORK_ERROR, 'x')));
    vi.mocked(createBrowserContext).mockResolvedValue(fakeManager());
    vi.mocked(ensureAuthenticated).mockRejectedValue(new Error('interaction required'));
    vi.mocked(closeBrowser).mockResolvedValue(undefined as never);

    const result = await refreshTokensViaBrowser();

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.code).toBe(ErrorCode.UNKNOWN);
      expect(result.error.message).toContain('interaction required');
    }
    expect(closeBrowser).toHaveBeenCalled();
  });

  it('ignores cleanup errors and still returns UNKNOWN on a non-Error throw', async () => {
    vi.mocked(extractSubstrateToken).mockReturnValueOnce(tokenInfo(60 * 60 * 1000));
    vi.mocked(refreshTokensViaHttp).mockResolvedValue(err(createError(ErrorCode.NETWORK_ERROR, 'x')));
    vi.mocked(createBrowserContext).mockResolvedValue(fakeManager());
    vi.mocked(ensureAuthenticated).mockRejectedValue('string failure');
    vi.mocked(closeBrowser).mockRejectedValue(new Error('cleanup failed'));

    const result = await refreshTokensViaBrowser();

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.code).toBe(ErrorCode.UNKNOWN);
      expect(result.error.message).toContain('Unknown error');
    }
  });

  it('returns UNKNOWN with manager null when createBrowserContext throws', async () => {
    vi.mocked(extractSubstrateToken).mockReturnValueOnce(tokenInfo(60 * 60 * 1000));
    vi.mocked(refreshTokensViaHttp).mockResolvedValue(err(createError(ErrorCode.NETWORK_ERROR, 'x')));
    vi.mocked(createBrowserContext).mockRejectedValue(new Error('no browser'));

    const result = await refreshTokensViaBrowser();

    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.code).toBe(ErrorCode.UNKNOWN);
    expect(closeBrowser).not.toHaveBeenCalled();
  });
});
