/**
 * Tests for browserless HTTP token refresh.
 * 
 * Tests the MSAL cache extraction and session state update logic.
 * Network calls are mocked to avoid hitting real Azure AD endpoints.
 */

import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';

// We need to mock the session-store and token-extractor modules
// before importing the module under test.
vi.mock('./session-store.js', () => ({
  readSessionState: vi.fn(),
  writeSessionState: vi.fn(),
  getTeamsOrigin: vi.fn(),
}));

vi.mock('./token-extractor.js', () => ({
  clearTokenCache: vi.fn(),
}));

import { refreshTokensViaHttp } from './token-refresh-http.js';
import { readSessionState, writeSessionState, getTeamsOrigin } from './session-store.js';
import { clearTokenCache } from './token-extractor.js';
import type { SessionState } from './session-store.js';

// ============================================================================
// Test Fixtures
// ============================================================================

const MOCK_HOME_ACCOUNT_ID = 'user-guid.tenant-guid';
const MOCK_CLIENT_ID = '5e3ce6c0-2b1f-4285-8d4b-75ee78787346';
const MOCK_TENANT_ID = 'tenant-guid';
const MOCK_ENVIRONMENT = 'login.windows.net';

/** Creates a minimal MSAL refresh token localStorage entry. */
function makeRefreshTokenEntry() {
  return {
    name: `${MOCK_HOME_ACCOUNT_ID}-${MOCK_ENVIRONMENT}-refreshtoken-${MOCK_CLIENT_ID}----`,
    value: JSON.stringify({
      credentialType: 'RefreshToken',
      homeAccountId: MOCK_HOME_ACCOUNT_ID,
      environment: MOCK_ENVIRONMENT,
      clientId: MOCK_CLIENT_ID,
      secret: 'mock-refresh-token-secret',
      expiresOn: String(Math.floor(Date.now() / 1000) + 86400),
      lastUpdatedAt: String(Date.now()),
    }),
  };
}

/** Creates a minimal MSAL access token localStorage entry. */
function makeAccessTokenEntry(resource: string, target: string, expiresInSeconds = 3600) {
  const now = Math.floor(Date.now() / 1000);
  return {
    name: `${MOCK_HOME_ACCOUNT_ID}-${MOCK_ENVIRONMENT}-accesstoken-${MOCK_CLIENT_ID}-${MOCK_TENANT_ID}-${target.toLowerCase()}`,
    value: JSON.stringify({
      credentialType: 'AccessToken',
      homeAccountId: MOCK_HOME_ACCOUNT_ID,
      environment: MOCK_ENVIRONMENT,
      clientId: MOCK_CLIENT_ID,
      realm: MOCK_TENANT_ID,
      target,
      tokenType: 'Bearer',
      secret: 'old-access-token',
      expiresOn: String(now + expiresInSeconds),
      extendedExpiresOn: String(now + expiresInSeconds + 3600),
      cachedAt: String(now),
    }),
  };
}

/** Creates a mock session state with MSAL cache entries. */
function makeMockSessionState(): SessionState {
  return {
    cookies: [
      {
        name: 'skypetoken_asm',
        value: 'old-skype-token',
        domain: '.asyncgw.teams.microsoft.com',
        path: '/',
        expires: Date.now() / 1000 + 3600,
        httpOnly: true,
        secure: true,
        sameSite: 'None',
      },
      {
        name: 'skypetoken_asm',
        value: 'old-skype-token',
        domain: '.asm.skype.com',
        path: '/',
        expires: Date.now() / 1000 + 3600,
        httpOnly: true,
        secure: true,
        sameSite: 'None',
      },
      {
        name: 'authtoken',
        value: 'Bearer%3Dold-auth-token',
        domain: 'teams.microsoft.com',
        path: '/',
        expires: Date.now() / 1000 + 3600,
        httpOnly: false,
        secure: true,
        sameSite: 'None',
      },
    ],
    origins: [
      {
        origin: 'https://teams.microsoft.com',
        localStorage: [
          makeRefreshTokenEntry(),
          makeAccessTokenEntry('substrate.office.com', 'https://substrate.office.com/.default'),
          makeAccessTokenEntry('api.spaces.skype.com', 'https://api.spaces.skype.com/.default'),
          makeAccessTokenEntry('chatsvcagg.teams.microsoft.com', 'https://chatsvcagg.teams.microsoft.com/.default'),
        ],
      },
    ],
  };
}

/** Creates a mock Azure AD token response. */
function makeTokenResponse(scope: string, expiresIn = 3600) {
  return {
    access_token: `new-access-token-for-${scope}`,
    refresh_token: 'new-refresh-token',
    token_type: 'Bearer',
    expires_in: expiresIn,
    scope,
    ext_expires_in: expiresIn + 3600,
  };
}

// ============================================================================
// Tests
// ============================================================================

describe('refreshTokensViaHttp', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    vi.stubGlobal('fetch', vi.fn());
  });

  afterEach(() => {
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
  });

  it('returns AUTH_REQUIRED when no session state exists', async () => {
    vi.mocked(readSessionState).mockReturnValue(null);

    const result = await refreshTokensViaHttp();

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.code).toBe('AUTH_REQUIRED');
      expect(result.error.message).toContain('No session state found');
    }
  });

  it('returns AUTH_REQUIRED when no refresh token in session', async () => {
    const state = makeMockSessionState();
    // Remove the refresh token entry
    const origin = state.origins[0];
    origin.localStorage = origin.localStorage.filter(item => {
      try {
        const val = JSON.parse(item.value);
        return val.credentialType !== 'RefreshToken';
      } catch { return true; }
    });

    vi.mocked(readSessionState).mockReturnValue(state);
    vi.mocked(getTeamsOrigin).mockReturnValue(state.origins[0]);

    const result = await refreshTokensViaHttp();

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.code).toBe('AUTH_REQUIRED');
      expect(result.error.message).toContain('No MSAL refresh token');
    }
  });

  it('refreshes all tokens successfully', async () => {
    const state = makeMockSessionState();
    vi.mocked(readSessionState).mockReturnValue(state);
    vi.mocked(getTeamsOrigin).mockReturnValue(state.origins[0]);

    // Mock Azure AD token endpoint responses
    vi.mocked(fetch).mockImplementation(async (url: RequestInfo | URL) => {
      const urlStr = String(url);
      if (urlStr.includes('login.microsoftonline.com')) {
        return new Response(JSON.stringify(makeTokenResponse('mocked-scope')), {
          status: 200,
          headers: { 'Content-Type': 'application/json' },
        });
      }
      if (urlStr.includes('authsvc.teams.microsoft.com')) {
        return new Response(JSON.stringify({
          tokens: { skypeToken: 'new-skype-token', expiresIn: 86400 },
        }), {
          status: 200,
          headers: { 'Content-Type': 'application/json' },
        });
      }
      return new Response('unexpected URL', { status: 500 });
    });

    const result = await refreshTokensViaHttp();

    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value.tokensRefreshed).toBe(4);
      expect(result.value.skypeTokenRefreshed).toBe(true);
      expect(result.value.refreshTokenRotated).toBe(true);
    }

    // Verify session state was written back
    expect(writeSessionState).toHaveBeenCalledOnce();

    // Verify token cache was cleared
    expect(clearTokenCache).toHaveBeenCalledOnce();

    // Verify fetch was called 5 times (4 token refreshes + 1 skype exchange)
    expect(vi.mocked(fetch)).toHaveBeenCalledTimes(5);
  });

  it('updates skypetoken_asm cookies in session state', async () => {
    const state = makeMockSessionState();
    vi.mocked(readSessionState).mockReturnValue(state);
    vi.mocked(getTeamsOrigin).mockReturnValue(state.origins[0]);

    vi.mocked(fetch).mockImplementation(async (url: RequestInfo | URL) => {
      const urlStr = String(url);
      if (urlStr.includes('login.microsoftonline.com')) {
        return new Response(JSON.stringify(makeTokenResponse('scope')), {
          status: 200, headers: { 'Content-Type': 'application/json' },
        });
      }
      if (urlStr.includes('authsvc.teams.microsoft.com')) {
        return new Response(JSON.stringify({
          tokens: { skypeToken: 'brand-new-skype-token', expiresIn: 86400 },
        }), {
          status: 200, headers: { 'Content-Type': 'application/json' },
        });
      }
      return new Response('', { status: 500 });
    });

    await refreshTokensViaHttp();

    // Check the session state that was written
    const writtenState = vi.mocked(writeSessionState).mock.calls[0][0];
    const skypeCookies = writtenState.cookies.filter(
      (c: { name: string }) => c.name === 'skypetoken_asm'
    );
    expect(skypeCookies.length).toBe(2);
    for (const cookie of skypeCookies) {
      expect(cookie.value).toBe('brand-new-skype-token');
    }
  });

  it('handles Azure AD error (expired refresh token)', async () => {
    const state = makeMockSessionState();
    vi.mocked(readSessionState).mockReturnValue(state);
    vi.mocked(getTeamsOrigin).mockReturnValue(state.origins[0]);

    vi.mocked(fetch).mockImplementation(async () => {
      return new Response(JSON.stringify({
        error: 'invalid_grant',
        error_description: 'AADSTS700082: The refresh token has expired.',
      }), { status: 400 });
    });

    const result = await refreshTokensViaHttp();

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.code).toBe('AUTH_EXPIRED');
      expect(result.error.message).toContain('refresh token has expired');
    }
  });

  it('continues with remaining scopes if one fails with network error', async () => {
    const state = makeMockSessionState();
    vi.mocked(readSessionState).mockReturnValue(state);
    vi.mocked(getTeamsOrigin).mockReturnValue(state.origins[0]);

    let callCount = 0;
    vi.mocked(fetch).mockImplementation(async (url: RequestInfo | URL) => {
      const urlStr = String(url);
      if (urlStr.includes('login.microsoftonline.com')) {
        callCount++;
        if (callCount === 1) {
          // First scope fails with network error
          throw new Error('ECONNRESET');
        }
        return new Response(JSON.stringify(makeTokenResponse('scope')), {
          status: 200, headers: { 'Content-Type': 'application/json' },
        });
      }
      if (urlStr.includes('authsvc.teams.microsoft.com')) {
        return new Response(JSON.stringify({
          tokens: { skypeToken: 'new-skype-token', expiresIn: 86400 },
        }), {
          status: 200, headers: { 'Content-Type': 'application/json' },
        });
      }
      return new Response('', { status: 500 });
    });

    const result = await refreshTokensViaHttp();

    expect(result.ok).toBe(true);
    if (result.ok) {
      // 3 of 4 scopes succeeded (first one failed with network error)
      expect(result.value.tokensRefreshed).toBe(3);
    }
  });

  it('persists already-refreshed tokens when a later scope returns AUTH_EXPIRED', async () => {
    // REFRESH_SCOPES order: substrate, api.spaces.skype.com, chatsvcagg, graph.
    // Simulate the first three scopes (substrate, api.spaces.skype.com,
    // chatsvcagg) succeeding, then Conditional Access / missing consent denying
    // graph.microsoft.com specifically with an auth error. The successful
    // refreshes for the earlier scopes must still be persisted.
    const state = makeMockSessionState();
    vi.mocked(readSessionState).mockReturnValue(state);
    vi.mocked(getTeamsOrigin).mockReturnValue(state.origins[0]);

    let callCount = 0;
    vi.mocked(fetch).mockImplementation(async (url: RequestInfo | URL) => {
      const urlStr = String(url);
      if (urlStr.includes('login.microsoftonline.com')) {
        callCount++;
        if (callCount === 4) {
          // Fourth scope (graph.microsoft.com) fails with an auth error
          return new Response(JSON.stringify({
            error: 'invalid_grant',
            error_description: 'AADSTS65001: The user or administrator has not consented.',
          }), { status: 400 });
        }
        return new Response(JSON.stringify(makeTokenResponse('scope')), {
          status: 200, headers: { 'Content-Type': 'application/json' },
        });
      }
      if (urlStr.includes('authsvc.teams.microsoft.com')) {
        return new Response(JSON.stringify({
          tokens: { skypeToken: 'new-skype-token', expiresIn: 86400 },
        }), {
          status: 200, headers: { 'Content-Type': 'application/json' },
        });
      }
      return new Response('', { status: 500 });
    });

    const result = await refreshTokensViaHttp();

    expect(result.ok).toBe(true);
    if (result.ok) {
      // 3 of 4 scopes succeeded (graph failed with an auth error)
      expect(result.value.tokensRefreshed).toBe(3);
    }
    // The successful refreshes must still be written to session state, not
    // discarded because one later scope hit an auth error.
    expect(writeSessionState).toHaveBeenCalled();
    expect(clearTokenCache).toHaveBeenCalled();
  });

  it('still refreshes later scopes when the first scope fails with a resource-specific AUTH_EXPIRED', async () => {
    // refreshAccessToken() classifies any 400/401 as AUTH_EXPIRED, which covers
    // both "the refresh token itself is dead" AND "missing consent/Conditional
    // Access for this one resource only". If the very first scope (substrate)
    // hits the latter, the remaining scopes must still be attempted rather than
    // aborting immediately, since the refresh token may well still be valid for
    // the other resources.
    const state = makeMockSessionState();
    vi.mocked(readSessionState).mockReturnValue(state);
    vi.mocked(getTeamsOrigin).mockReturnValue(state.origins[0]);

    let callCount = 0;
    vi.mocked(fetch).mockImplementation(async (url: RequestInfo | URL) => {
      const urlStr = String(url);
      if (urlStr.includes('login.microsoftonline.com')) {
        callCount++;
        if (callCount === 1) {
          // First scope (substrate) fails with a resource-specific auth error,
          // not necessarily a dead refresh token.
          return new Response(JSON.stringify({
            error: 'invalid_grant',
            error_description: 'AADSTS65001: The user or administrator has not consented.',
          }), { status: 400 });
        }
        return new Response(JSON.stringify(makeTokenResponse('scope')), {
          status: 200, headers: { 'Content-Type': 'application/json' },
        });
      }
      if (urlStr.includes('authsvc.teams.microsoft.com')) {
        return new Response(JSON.stringify({
          tokens: { skypeToken: 'new-skype-token', expiresIn: 86400 },
        }), {
          status: 200, headers: { 'Content-Type': 'application/json' },
        });
      }
      return new Response('', { status: 500 });
    });

    const result = await refreshTokensViaHttp();

    expect(result.ok).toBe(true);
    if (result.ok) {
      // 3 of 4 scopes succeeded; only substrate failed
      expect(result.value.tokensRefreshed).toBe(3);
    }
    expect(writeSessionState).toHaveBeenCalled();
  });

  it('handles skype token exchange failure gracefully', async () => {
    const state = makeMockSessionState();
    vi.mocked(readSessionState).mockReturnValue(state);
    vi.mocked(getTeamsOrigin).mockReturnValue(state.origins[0]);

    vi.mocked(fetch).mockImplementation(async (url: RequestInfo | URL) => {
      const urlStr = String(url);
      if (urlStr.includes('login.microsoftonline.com')) {
        return new Response(JSON.stringify(makeTokenResponse('scope')), {
          status: 200, headers: { 'Content-Type': 'application/json' },
        });
      }
      if (urlStr.includes('authsvc.teams.microsoft.com')) {
        return new Response('Forbidden', { status: 403 });
      }
      return new Response('', { status: 500 });
    });

    const result = await refreshTokensViaHttp();

    // Should still succeed — skype token failure is non-fatal
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value.tokensRefreshed).toBe(4);
      expect(result.value.skypeTokenRefreshed).toBe(false);
    }
  });

  it('updates MSAL access token cache entries in localStorage', async () => {
    const state = makeMockSessionState();
    vi.mocked(readSessionState).mockReturnValue(state);
    vi.mocked(getTeamsOrigin).mockReturnValue(state.origins[0]);

    vi.mocked(fetch).mockImplementation(async (url: RequestInfo | URL) => {
      const urlStr = String(url);
      if (urlStr.includes('login.microsoftonline.com')) {
        return new Response(JSON.stringify({
          access_token: 'fresh-substrate-token',
          refresh_token: 'new-rt',
          token_type: 'Bearer',
          expires_in: 7200,
          scope: 'https://substrate.office.com/.default',
          ext_expires_in: 10800,
        }), {
          status: 200, headers: { 'Content-Type': 'application/json' },
        });
      }
      if (urlStr.includes('authsvc.teams.microsoft.com')) {
        return new Response(JSON.stringify({
          tokens: { skypeToken: 'st', expiresIn: 86400 },
        }), {
          status: 200, headers: { 'Content-Type': 'application/json' },
        });
      }
      return new Response('', { status: 500 });
    });

    await refreshTokensViaHttp();

    // Check that the Substrate access token was updated in localStorage
    const writtenState = vi.mocked(writeSessionState).mock.calls[0][0];
    const origin = writtenState.origins[0];
    const substrateEntry = origin.localStorage.find((item: { name: string; value: string }) => {
      try {
        const val = JSON.parse(item.value);
        return val.credentialType === 'AccessToken' && val.target?.includes('substrate.office.com');
      } catch { return false; }
    });

    expect(substrateEntry).toBeDefined();
    const parsed = JSON.parse(substrateEntry!.value);
    expect(parsed.secret).toBe('fresh-substrate-token');
  });

  it('updates refresh token when Azure AD rotates it', async () => {
    const state = makeMockSessionState();
    vi.mocked(readSessionState).mockReturnValue(state);
    vi.mocked(getTeamsOrigin).mockReturnValue(state.origins[0]);

    vi.mocked(fetch).mockImplementation(async (url: RequestInfo | URL) => {
      const urlStr = String(url);
      if (urlStr.includes('login.microsoftonline.com')) {
        return new Response(JSON.stringify({
          access_token: 'new-at',
          refresh_token: 'rotated-refresh-token',
          token_type: 'Bearer',
          expires_in: 3600,
          scope: 'scope',
        }), {
          status: 200, headers: { 'Content-Type': 'application/json' },
        });
      }
      if (urlStr.includes('authsvc.teams.microsoft.com')) {
        return new Response(JSON.stringify({
          tokens: { skypeToken: 'st', expiresIn: 86400 },
        }), {
          status: 200, headers: { 'Content-Type': 'application/json' },
        });
      }
      return new Response('', { status: 500 });
    });

    await refreshTokensViaHttp();

    const writtenState = vi.mocked(writeSessionState).mock.calls[0][0];
    const origin = writtenState.origins[0];
    const rtEntry = origin.localStorage.find((item: { name: string; value: string }) => {
      try {
        const val = JSON.parse(item.value);
        return val.credentialType === 'RefreshToken';
      } catch { return false; }
    });

    expect(rtEntry).toBeDefined();
    const parsed = JSON.parse(rtEntry!.value);
    expect(parsed.secret).toBe('rotated-refresh-token');
  });
});

// ============================================================================
// Additional branch and function coverage
// ============================================================================

describe('refreshTokensViaHttp - additional branch and function coverage', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    vi.stubGlobal('fetch', vi.fn());
  });

  afterEach(() => {
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
  });

  /** Builds a fetch implementation that succeeds for all token scopes and returns the given skype tokens. */
  function successFetch(skypeTokens: { skypeToken?: string; expiresIn?: number } = { skypeToken: 'new-skype-token', expiresIn: 86400 }) {
    return async (url: RequestInfo | URL) => {
      const urlStr = String(url);
      if (urlStr.includes('login.microsoftonline.com')) {
        return new Response(JSON.stringify(makeTokenResponse('scope')), {
          status: 200, headers: { 'Content-Type': 'application/json' },
        });
      }
      if (urlStr.includes('authsvc.teams.microsoft.com')) {
        return new Response(JSON.stringify({ tokens: skypeTokens }), {
          status: 200, headers: { 'Content-Type': 'application/json' },
        });
      }
      return new Response('', { status: 500 });
    };
  }

  it('returns AUTH_REQUIRED when the teams origin has no localStorage', async () => {
    vi.mocked(readSessionState).mockReturnValue(makeMockSessionState());
    vi.mocked(getTeamsOrigin).mockReturnValue(null);

    const result = await refreshTokensViaHttp();

    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.code).toBe('AUTH_REQUIRED');
  });

  it('logs the client id when extraction fails because tenant id is missing', async () => {
    const state = makeMockSessionState();
    // Keep the refresh token entry but drop all access tokens so no realm/tenantId is found.
    state.origins[0].localStorage = state.origins[0].localStorage.filter((item) => {
      try { return JSON.parse(item.value).credentialType !== 'AccessToken'; }
      catch { return true; }
    });
    vi.mocked(readSessionState).mockReturnValue(state);
    vi.mocked(getTeamsOrigin).mockReturnValue(state.origins[0]);

    const result = await refreshTokensViaHttp();

    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.code).toBe('AUTH_REQUIRED');
  });

  it('skips invalid-JSON localStorage entries during extraction and cache update', async () => {
    const state = makeMockSessionState();
    state.origins[0].localStorage.push({ name: 'garbage', value: '{not valid json' });
    vi.mocked(readSessionState).mockReturnValue(state);
    vi.mocked(getTeamsOrigin).mockReturnValue(state.origins[0]);
    vi.mocked(fetch).mockImplementation(successFetch());

    const result = await refreshTokensViaHttp();

    expect(result.ok).toBe(true);
  });

  it('returns AUTH_REQUIRED when the teams origin disappears after extraction', async () => {
    const state = makeMockSessionState();
    vi.mocked(readSessionState).mockReturnValue(state);
    vi.mocked(getTeamsOrigin)
      .mockReturnValueOnce(state.origins[0])
      .mockReturnValueOnce(null);

    const result = await refreshTokensViaHttp();

    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.code).toBe('AUTH_REQUIRED');
  });

  it('returns UNKNOWN when every scope fails with a non-auth (500) error', async () => {
    const state = makeMockSessionState();
    vi.mocked(readSessionState).mockReturnValue(state);
    vi.mocked(getTeamsOrigin).mockReturnValue(state.origins[0]);
    vi.mocked(fetch).mockImplementation(async (url: RequestInfo | URL) => {
      if (String(url).includes('login.microsoftonline.com')) {
        return new Response('boom-not-json', { status: 500 });
      }
      return new Response('', { status: 500 });
    });

    const result = await refreshTokensViaHttp();

    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.code).toBe('UNKNOWN');
  });

  it('returns UNKNOWN when every scope rejects with a non-Error value', async () => {
    const state = makeMockSessionState();
    vi.mocked(readSessionState).mockReturnValue(state);
    vi.mocked(getTeamsOrigin).mockReturnValue(state.origins[0]);
    vi.mocked(fetch).mockImplementation(async (url: RequestInfo | URL) => {
      if (String(url).includes('login.microsoftonline.com')) {
        throw 'string failure';
      }
      return new Response('', { status: 500 });
    });

    const result = await refreshTokensViaHttp();

    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.code).toBe('UNKNOWN');
  });

  it('returns UNKNOWN when every scope aborts with an AbortError', async () => {
    const state = makeMockSessionState();
    vi.mocked(readSessionState).mockReturnValue(state);
    vi.mocked(getTeamsOrigin).mockReturnValue(state.origins[0]);
    vi.mocked(fetch).mockImplementation(async (url: RequestInfo | URL) => {
      if (String(url).includes('login.microsoftonline.com')) {
        const error = new Error('aborted');
        error.name = 'AbortError';
        throw error;
      }
      return new Response('', { status: 500 });
    });

    const result = await refreshTokensViaHttp();

    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.code).toBe('UNKNOWN');
  });

  it('treats an AAD error without error_description as AUTH_EXPIRED', async () => {
    const state = makeMockSessionState();
    vi.mocked(readSessionState).mockReturnValue(state);
    vi.mocked(getTeamsOrigin).mockReturnValue(state.origins[0]);
    vi.mocked(fetch).mockImplementation(async () =>
      new Response(JSON.stringify({ error: 'invalid_grant' }), { status: 400 })
    );

    const result = await refreshTokensViaHttp();

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.code).toBe('AUTH_EXPIRED');
      expect(result.error.message).toContain('invalid_grant');
    }
  });

  it('handles a token endpoint whose error body cannot be read', async () => {
    const state = makeMockSessionState();
    vi.mocked(readSessionState).mockReturnValue(state);
    vi.mocked(getTeamsOrigin).mockReturnValue(state.origins[0]);
    vi.mocked(fetch).mockImplementation(async (url: RequestInfo | URL) => {
      if (String(url).includes('login.microsoftonline.com')) {
        return { ok: false, status: 500, text: () => Promise.reject(new Error('unreadable')) } as unknown as Response;
      }
      return new Response('', { status: 500 });
    });

    const result = await refreshTokensViaHttp();

    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.code).toBe('UNKNOWN');
  });

  it('creates a new cache entry and defaults token_type when the response omits it', async () => {
    const state = makeMockSessionState();
    vi.mocked(readSessionState).mockReturnValue(state);
    vi.mocked(getTeamsOrigin).mockReturnValue(state.origins[0]);
    vi.mocked(fetch).mockImplementation(async (url: RequestInfo | URL) => {
      if (String(url).includes('login.microsoftonline.com')) {
        return new Response(JSON.stringify({
          access_token: 'at-no-type',
          expires_in: 3600,
          scope: 'https://graph.microsoft.com/.default',
          // no token_type, no ext_expires_in, no refresh_token
        }), { status: 200, headers: { 'Content-Type': 'application/json' } });
      }
      if (String(url).includes('authsvc.teams.microsoft.com')) {
        return new Response(JSON.stringify({ tokens: { skypeToken: 'st', expiresIn: 86400 } }), {
          status: 200, headers: { 'Content-Type': 'application/json' },
        });
      }
      return new Response('', { status: 500 });
    });

    const result = await refreshTokensViaHttp();

    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value.refreshTokenRotated).toBe(false);
  });

  it('pushes new cookies when skype/auth cookies are not already present', async () => {
    const state = makeMockSessionState();
    state.cookies = []; // force the "push new cookie" branches
    vi.mocked(readSessionState).mockReturnValue(state);
    vi.mocked(getTeamsOrigin).mockReturnValue(state.origins[0]);
    vi.mocked(fetch).mockImplementation(successFetch());

    const result = await refreshTokensViaHttp();

    expect(result.ok).toBe(true);
    const written = vi.mocked(writeSessionState).mock.calls[0][0];
    expect(written.cookies.some((c: { name: string }) => c.name === 'skypetoken_asm')).toBe(true);
    expect(written.cookies.some((c: { name: string }) => c.name === 'authtoken')).toBe(true);
  });

  it('defaults the auth-token expiry when the spaces token omits expires_in', async () => {
    const state = makeMockSessionState();
    vi.mocked(readSessionState).mockReturnValue(state);
    vi.mocked(getTeamsOrigin).mockReturnValue(state.origins[0]);
    vi.mocked(fetch).mockImplementation(async (url: RequestInfo | URL) => {
      if (String(url).includes('login.microsoftonline.com')) {
        return new Response(JSON.stringify({
          access_token: 'spaces-at',
          token_type: 'Bearer',
          scope: 'https://api.spaces.skype.com/.default',
          // expires_in omitted -> exercises the `?? 3600` / `?? expires_in` fallbacks
        }), { status: 200, headers: { 'Content-Type': 'application/json' } });
      }
      if (String(url).includes('authsvc.teams.microsoft.com')) {
        return new Response(JSON.stringify({ tokens: { skypeToken: 'st' } }), {
          status: 200, headers: { 'Content-Type': 'application/json' },
        });
      }
      return new Response('', { status: 500 });
    });

    const result = await refreshTokensViaHttp();

    expect(result.ok).toBe(true);
  });

  it('treats a skype exchange that returns no token as non-fatal', async () => {
    const state = makeMockSessionState();
    vi.mocked(readSessionState).mockReturnValue(state);
    vi.mocked(getTeamsOrigin).mockReturnValue(state.origins[0]);
    vi.mocked(fetch).mockImplementation(successFetch({})); // tokens: {} -> no skypeToken

    const result = await refreshTokensViaHttp();

    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value.skypeTokenRefreshed).toBe(false);
  });

  it('defaults the skype expiry when expiresIn is omitted', async () => {
    const state = makeMockSessionState();
    vi.mocked(readSessionState).mockReturnValue(state);
    vi.mocked(getTeamsOrigin).mockReturnValue(state.origins[0]);
    vi.mocked(fetch).mockImplementation(successFetch({ skypeToken: 'only-token' })); // no expiresIn

    const result = await refreshTokensViaHttp();

    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value.skypeTokenRefreshed).toBe(true);
  });

  it('treats an unreadable skype error response as non-fatal', async () => {
    const state = makeMockSessionState();
    vi.mocked(readSessionState).mockReturnValue(state);
    vi.mocked(getTeamsOrigin).mockReturnValue(state.origins[0]);
    vi.mocked(fetch).mockImplementation(async (url: RequestInfo | URL) => {
      if (String(url).includes('login.microsoftonline.com')) {
        return new Response(JSON.stringify(makeTokenResponse('scope')), {
          status: 200, headers: { 'Content-Type': 'application/json' },
        });
      }
      return { ok: false, status: 403, text: () => Promise.reject(new Error('nope')) } as unknown as Response;
    });

    const result = await refreshTokensViaHttp();

    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value.skypeTokenRefreshed).toBe(false);
  });

  it('treats a skype exchange that rejects with a non-Error as non-fatal', async () => {
    const state = makeMockSessionState();
    vi.mocked(readSessionState).mockReturnValue(state);
    vi.mocked(getTeamsOrigin).mockReturnValue(state.origins[0]);
    vi.mocked(fetch).mockImplementation(async (url: RequestInfo | URL) => {
      if (String(url).includes('login.microsoftonline.com')) {
        return new Response(JSON.stringify(makeTokenResponse('scope')), {
          status: 200, headers: { 'Content-Type': 'application/json' },
        });
      }
      throw 'skype string failure';
    });

    const result = await refreshTokensViaHttp();

    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value.skypeTokenRefreshed).toBe(false);
  });

  it('treats a skype exchange timeout as non-fatal and fires the abort timer', async () => {
    vi.useFakeTimers();
    try {
      const state = makeMockSessionState();
      vi.mocked(readSessionState).mockReturnValue(state);
      vi.mocked(getTeamsOrigin).mockReturnValue(state.origins[0]);
      vi.mocked(fetch).mockImplementation((url: RequestInfo | URL, init?: RequestInit) => {
        if (String(url).includes('login.microsoftonline.com')) {
          return Promise.resolve(new Response(JSON.stringify(makeTokenResponse('scope')), {
            status: 200, headers: { 'Content-Type': 'application/json' },
          }));
        }
        // authsvc hangs until its abort signal fires (driven by the real setTimeout)
        return new Promise<Response>((_resolve, reject) => {
          init?.signal?.addEventListener('abort', () => {
            const error = new Error('aborted');
            error.name = 'AbortError';
            reject(error);
          });
        });
      });

      const promise = refreshTokensViaHttp();
      await vi.advanceTimersByTimeAsync(20000);
      const result = await promise;

      expect(result.ok).toBe(true);
      if (result.ok) expect(result.value.skypeTokenRefreshed).toBe(false);
    } finally {
      vi.useRealTimers();
    }
  });
});
