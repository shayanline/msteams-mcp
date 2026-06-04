/**
 * Unit tests for auth tools (real handlers).
 */
import { describe, it, expect, vi, beforeEach } from 'vitest';
import './index.js';

vi.mock('../auth/session-store.js', () => ({
  hasSessionState: vi.fn(),
  isSessionLikelyExpired: vi.fn(),
  clearSessionState: vi.fn(),
}));
vi.mock('../auth/token-extractor.js', () => ({
  getSubstrateTokenStatus: vi.fn(),
  getMessageAuthStatus: vi.fn(),
  extractMessageAuth: vi.fn(),
  extractCsaToken: vi.fn(),
  clearTokenCache: vi.fn(),
}));
vi.mock('../browser/context.js', () => ({
  createBrowserContext: vi.fn(),
  closeBrowser: vi.fn(),
}));
vi.mock('../browser/auth.js', () => ({
  ensureAuthenticated: vi.fn(),
  forceNewLogin: vi.fn(),
  getAuthStatus: vi.fn(),
}));
vi.mock('../utils/logger.js', () => ({ info: vi.fn(), warn: vi.fn(), error: vi.fn() }));

import {
  hasSessionState,
  isSessionLikelyExpired,
  clearSessionState,
} from '../auth/session-store.js';
import {
  getSubstrateTokenStatus,
  getMessageAuthStatus,
  extractMessageAuth,
  extractCsaToken,
  clearTokenCache,
} from '../auth/token-extractor.js';
import { createBrowserContext, closeBrowser } from '../browser/context.js';
import { ensureAuthenticated, forceNewLogin, getAuthStatus } from '../browser/auth.js';
import { loginTool, statusTool, LoginInputSchema } from './auth-tools.js';

function makeManager() {
  return {
    page: {},
    context: { clearCookies: vi.fn().mockResolvedValue(undefined) },
  } as never;
}

function makeServer(overrides: Record<string, unknown> = {}) {
  return {
    server: {
      getBrowserManager: vi.fn().mockReturnValue(null),
      resetBrowserState: vi.fn(),
      setBrowserManager: vi.fn(),
      markInitialised: vi.fn(),
      isInitialisedState: vi.fn().mockReturnValue(false),
      ensureBrowser: vi.fn(),
      ...overrides,
    },
  } as never;
}

beforeEach(() => {
  vi.clearAllMocks();
  vi.mocked(closeBrowser).mockResolvedValue(undefined as never);
  vi.mocked(createBrowserContext).mockResolvedValue(makeManager());
  vi.mocked(ensureAuthenticated).mockResolvedValue(undefined as never);
  vi.mocked(forceNewLogin).mockResolvedValue(undefined as never);
});

describe('LoginInputSchema', () => {
  it('defaults forceNew to false', () => {
    expect(LoginInputSchema.parse({}).forceNew).toBe(false);
  });
});

describe('loginTool', () => {
  it('returns early when a valid token already exists (fast path)', async () => {
    vi.mocked(getSubstrateTokenStatus).mockReturnValue({
      hasToken: true, minutesRemaining: 30, expiresAt: 'soon',
    } as never);
    const ctx = makeServer();
    const res = await loginTool.handler({ forceNew: false }, ctx);
    expect(res.success).toBe(true);
    if (res.success) expect(String(res.data.message)).toContain('Already authenticated');
    expect(createBrowserContext).not.toHaveBeenCalled();
  });

  it('returns early-skip when token below threshold then succeeds headless', async () => {
    vi.mocked(getSubstrateTokenStatus).mockReturnValue({
      hasToken: true, minutesRemaining: 2, expiresAt: 'soon',
    } as never);
    const ctx = makeServer();
    const res = await loginTool.handler({ forceNew: false }, ctx);
    expect(res.success).toBe(true);
    if (res.success) expect(String(res.data.message)).toContain('silently via SSO');
  });

  it('closes existing browser then succeeds via headless SSO', async () => {
    const existing = makeManager();
    vi.mocked(getSubstrateTokenStatus).mockReturnValue({ hasToken: false } as never);
    const ctx = makeServer({ getBrowserManager: vi.fn().mockReturnValue(existing) });
    const res = await loginTool.handler({ forceNew: false }, ctx);
    expect(res.success).toBe(true);
    if (res.success) expect(String(res.data.message)).toContain('silently via SSO');
    expect(closeBrowser).toHaveBeenCalled();
  });

  it('forceNew: clears state, headless fails, falls back to visible forceNewLogin', async () => {
    vi.mocked(getSubstrateTokenStatus).mockReturnValue({ hasToken: false } as never);
    // headless attempt rejects, cleanup closeBrowser rejects (covers inner catch),
    // then visible browser closeBrowser resolves.
    vi.mocked(ensureAuthenticated).mockRejectedValueOnce(new Error('needs interaction'));
    vi.mocked(closeBrowser).mockRejectedValueOnce(new Error('cleanup failed'));
    const ctx = makeServer();
    const res = await loginTool.handler({ forceNew: true }, ctx);
    expect(res.success).toBe(true);
    if (res.success) expect(String(res.data.message)).toContain('completed successfully');
    expect(clearSessionState).toHaveBeenCalled();
    expect(clearTokenCache).toHaveBeenCalled();
    expect(forceNewLogin).toHaveBeenCalled();
  });

  it('non-forceNew: headless fails, falls back to visible ensureAuthenticated', async () => {
    vi.mocked(getSubstrateTokenStatus).mockReturnValue({ hasToken: false } as never);
    vi.mocked(ensureAuthenticated)
      .mockRejectedValueOnce(new Error('needs interaction'))
      .mockResolvedValueOnce(undefined as never);
    const ctx = makeServer();
    const res = await loginTool.handler({ forceNew: false }, ctx);
    expect(res.success).toBe(true);
    if (res.success) expect(String(res.data.message)).toContain('completed successfully');
    expect(forceNewLogin).not.toHaveBeenCalled();
  });
});

describe('statusTool', () => {
  it('reports status with no running browser', async () => {
    vi.mocked(hasSessionState).mockReturnValue(true);
    vi.mocked(isSessionLikelyExpired).mockReturnValue(false);
    vi.mocked(getSubstrateTokenStatus).mockReturnValue({ hasToken: true, expiresAt: 'x', minutesRemaining: 5 } as never);
    vi.mocked(getMessageAuthStatus).mockReturnValue({ hasToken: false } as never);
    vi.mocked(extractMessageAuth).mockReturnValue(null as never);
    vi.mocked(extractCsaToken).mockReturnValue(null as never);
    const ctx = makeServer();
    const res = await statusTool.handler({}, ctx);
    expect(res.success).toBe(true);
    if (res.success) {
      expect((res.data.browser as Record<string, unknown>).running).toBe(false);
      expect((res.data.favorites as Record<string, unknown>).available).toBe(false);
      expect(res.data.authentication).toBeNull();
    }
    expect(getAuthStatus).not.toHaveBeenCalled();
  });

  it('reports auth status when browser is running and initialised', async () => {
    vi.mocked(hasSessionState).mockReturnValue(true);
    vi.mocked(isSessionLikelyExpired).mockReturnValue(false);
    vi.mocked(getSubstrateTokenStatus).mockReturnValue({ hasToken: true, expiresAt: 'x', minutesRemaining: 50 } as never);
    vi.mocked(getMessageAuthStatus).mockReturnValue({ hasToken: true, expiresAt: 'y', minutesRemaining: 40 } as never);
    vi.mocked(extractMessageAuth).mockReturnValue({ token: 'm' } as never);
    vi.mocked(extractCsaToken).mockReturnValue({ token: 'c' } as never);
    vi.mocked(getAuthStatus).mockResolvedValue({ authenticated: true } as never);
    const ctx = makeServer({
      getBrowserManager: vi.fn().mockReturnValue(makeManager()),
      isInitialisedState: vi.fn().mockReturnValue(true),
    });
    const res = await statusTool.handler({}, ctx);
    expect(res.success).toBe(true);
    if (res.success) {
      expect((res.data.favorites as Record<string, unknown>).available).toBe(true);
      expect(res.data.authentication).toEqual({ authenticated: true });
    }
    expect(getAuthStatus).toHaveBeenCalled();
  });
});

describe('loginTool - log callbacks and error shapes', () => {
  it('invokes the headless progress callback during silent SSO', async () => {
    vi.mocked(getSubstrateTokenStatus).mockReturnValue({ hasToken: false } as never);
    vi.mocked(ensureAuthenticated).mockImplementationOnce(
      async (_page, _ctx, logCb) => {
        (logCb as (m: string) => void)?.('headless progress');
      }
    );
    const ctx = makeServer();
    const res = await loginTool.handler({ forceNew: false }, ctx);
    expect(res.success).toBe(true);
    if (res.success) expect(String(res.data.message)).toContain('silently via SSO');
  });

  it('invokes the visible login callback when falling back (non-forceNew)', async () => {
    vi.mocked(getSubstrateTokenStatus).mockReturnValue({ hasToken: false } as never);
    vi.mocked(ensureAuthenticated)
      .mockRejectedValueOnce(new Error('needs interaction'))
      .mockImplementationOnce(async (_page, _ctx, logCb) => {
        (logCb as (m: string) => void)?.('visible progress');
      });
    const ctx = makeServer();
    const res = await loginTool.handler({ forceNew: false }, ctx);
    expect(res.success).toBe(true);
    if (res.success) expect(String(res.data.message)).toContain('completed successfully');
  });

  it('invokes the forceNewLogin callback when forcing a fresh login', async () => {
    vi.mocked(getSubstrateTokenStatus).mockReturnValue({ hasToken: false } as never);
    vi.mocked(ensureAuthenticated).mockRejectedValueOnce(new Error('needs interaction'));
    vi.mocked(forceNewLogin).mockImplementationOnce(
      async (_page, _ctx, logCb) => {
        (logCb as (m: string) => void)?.('force progress');
      }
    );
    const ctx = makeServer();
    const res = await loginTool.handler({ forceNew: true }, ctx);
    expect(res.success).toBe(true);
    expect(forceNewLogin).toHaveBeenCalled();
  });

  it('handles a non-Error thrown by the headless attempt', async () => {
    vi.mocked(getSubstrateTokenStatus).mockReturnValue({ hasToken: false } as never);
    // Reject with a non-Error value to exercise the String(error) branch.
    vi.mocked(ensureAuthenticated).mockRejectedValueOnce('string failure');
    const ctx = makeServer();
    const res = await loginTool.handler({ forceNew: false }, ctx);
    expect(res.success).toBe(true);
    if (res.success) expect(String(res.data.message)).toContain('completed successfully');
  });
});
