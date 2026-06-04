/**
 * Unit tests for token-extractor.
 *
 * Covers jwtGrantsScope plus the full token/profile/config extraction surface.
 * Session IO (readSessionState/readTokenCache/...) is mocked; the pure
 * getTeamsOrigin helper is kept real via importActual so localStorage lookups
 * behave exactly as in production.
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';

vi.mock('./session-store.js', async (importActual) => {
  const actual = await importActual<typeof import('./session-store.js')>();
  return {
    ...actual,
    readSessionState: vi.fn(),
    readTokenCache: vi.fn(),
    writeTokenCache: vi.fn(),
    clearTokenCache: vi.fn(),
  };
});

import {
  jwtGrantsScope,
  extractSubstrateToken,
  getValidSubstrateToken,
  hasValidSubstrateToken,
  getSubstrateTokenStatus,
  extractTeamsToken,
  extractSkypeSpacesToken,
  extractGraphToken,
  extractRegionConfig,
  extractUserDetails,
  extractMessageAuth,
  getMessageAuthStatus,
  extractCsaToken,
  getUserProfile,
  getUserDisplayName,
  areTokensExpired,
  discoverConfig,
} from './token-extractor.js';
import {
  readSessionState,
  readTokenCache,
  writeTokenCache,
  type SessionState,
} from './session-store.js';

// ============================================================================
// Helpers
// ============================================================================

const b64 = (s: string) => Buffer.from(s).toString('base64');

/** Builds a fake (unsigned) JWT whose payload decodes to `payload`. */
function makeJwt(payload: Record<string, unknown>): string {
  return `${b64('{"alg":"none","typ":"JWT"}')}.${b64(JSON.stringify(payload))}.sig`;
}

const HOUR = 3600;
const futureExp = (secondsAhead = HOUR) => Math.floor(Date.now() / 1000) + secondsAhead;
const pastExp = () => Math.floor(Date.now() / 1000) - HOUR;

/** localStorage entry helper. */
const entry = (name: string, value: unknown) => ({
  name,
  value: typeof value === 'string' ? value : JSON.stringify(value),
});

/** Builds a session state with a single Teams origin holding the given entries. */
function sessionWith(
  localStorage: Array<{ name: string; value: string }>,
  cookies: SessionState['cookies'] = [],
  origin = 'https://teams.microsoft.com',
): SessionState {
  return {
    cookies,
    origins: [{ origin, localStorage }],
  };
}

beforeEach(() => {
  vi.clearAllMocks();
});

// ============================================================================
// jwtGrantsScope (existing coverage, preserved)
// ============================================================================

describe('jwtGrantsScope', () => {
  it('returns true when a delegated scope is present in scp', () => {
    const payload = { scp: 'Calendars.Read Calendars.ReadWrite Mail.Read' };
    expect(jwtGrantsScope(payload, 'Calendars.ReadWrite')).toBe(true);
  });

  it('returns false when the scope is absent from scp', () => {
    const payload = { scp: 'Calendars.Read Mail.Read' };
    expect(jwtGrantsScope(payload, 'Calendars.ReadWrite')).toBe(false);
  });

  it('matches scopes case-insensitively', () => {
    const payload = { scp: 'calendars.readwrite' };
    expect(jwtGrantsScope(payload, 'Calendars.ReadWrite')).toBe(true);
  });

  it('checks app roles when scp is absent', () => {
    const payload = { roles: ['Calendars.ReadWrite', 'Mail.Read'] };
    expect(jwtGrantsScope(payload, 'Calendars.ReadWrite')).toBe(true);
    expect(jwtGrantsScope(payload, 'User.Read.All')).toBe(false);
  });

  it('returns null when scopes cannot be determined', () => {
    expect(jwtGrantsScope({ aud: 'https://graph.microsoft.com' }, 'Calendars.ReadWrite')).toBeNull();
    expect(jwtGrantsScope(null, 'Calendars.ReadWrite')).toBeNull();
  });
});

// ============================================================================
// extractSubstrateToken
// ============================================================================

describe('extractSubstrateToken', () => {
  it('returns null when no session state is available', () => {
    vi.mocked(readSessionState).mockReturnValue(null);
    expect(extractSubstrateToken()).toBeNull();
  });

  it('extracts a valid substrate token (old format target)', () => {
    const token = makeJwt({ exp: futureExp(), oid: 'user-oid' });
    const state = sessionWith([
      entry('k', { target: 'https://substrate.office.com/search/SubstrateSearch', secret: token }),
    ]);
    const result = extractSubstrateToken(state);
    expect(result?.token).toBe(token);
    expect(result?.expiry).toBeInstanceOf(Date);
  });

  it('picks the token with the longest remaining validity', () => {
    const soon = makeJwt({ exp: futureExp(HOUR) });
    const later = makeJwt({ exp: futureExp(HOUR * 5) });
    const state = sessionWith([
      entry('a', { target: 'substrate.office.com/SubstrateSearch-Internal.ReadWrite', secret: soon }),
      entry('b', { target: 'substrate.office.com/SubstrateSearch-Internal.ReadWrite', secret: later }),
    ]);
    expect(extractSubstrateToken(state)?.token).toBe(later);
  });

  it('skips unrelated, non-jwt, expired and unparseable entries', () => {
    const state = sessionWith([
      entry('bad-json', 'not-json{'),
      entry('wrong-host', { target: 'https://graph.microsoft.com', secret: makeJwt({ exp: futureExp() }) }),
      entry('no-search', { target: 'https://substrate.office.com/other', secret: makeJwt({ exp: futureExp() }) }),
      entry('not-jwt', { target: 'substrate.office.com/SubstrateSearch', secret: 'plain' }),
      entry('no-exp', { target: 'substrate.office.com/SubstrateSearch', secret: makeJwt({ sub: 'x' }) }),
      entry('expired', { target: 'substrate.office.com/SubstrateSearch', secret: makeJwt({ exp: pastExp() }) }),
    ]);
    expect(extractSubstrateToken(state)).toBeNull();
  });

  it('returns null when the Teams origin has no localStorage', () => {
    const state: SessionState = { cookies: [], origins: [{ origin: 'https://example.com', localStorage: [] }] };
    expect(extractSubstrateToken(state)).toBeNull();
  });
});

// ============================================================================
// Cached substrate token access
// ============================================================================

describe('getValidSubstrateToken / hasValidSubstrateToken', () => {
  it('returns cached token when cache is still valid', () => {
    vi.mocked(readTokenCache).mockReturnValue({
      substrateToken: 'cached-token',
      substrateTokenExpiry: Date.now() + 60_000,
      extractedAt: Date.now(),
    });
    expect(getValidSubstrateToken()).toBe('cached-token');
    expect(readSessionState).not.toHaveBeenCalled();
  });

  it('extracts and caches when cache is missing', () => {
    vi.mocked(readTokenCache).mockReturnValue(null);
    const token = makeJwt({ exp: futureExp() });
    vi.mocked(readSessionState).mockReturnValue(
      sessionWith([entry('k', { target: 'substrate.office.com/SubstrateSearch', secret: token })]),
    );
    expect(getValidSubstrateToken()).toBe(token);
    expect(writeTokenCache).toHaveBeenCalledOnce();
  });

  it('returns null when no token can be extracted', () => {
    vi.mocked(readTokenCache).mockReturnValue(null);
    vi.mocked(readSessionState).mockReturnValue(null);
    expect(getValidSubstrateToken()).toBeNull();
    expect(hasValidSubstrateToken()).toBe(false);
  });

  it('hasValidSubstrateToken returns true when a token is available', () => {
    vi.mocked(readTokenCache).mockReturnValue({
      substrateToken: 'cached-token',
      substrateTokenExpiry: Date.now() + 60_000,
      extractedAt: Date.now(),
    });
    expect(hasValidSubstrateToken()).toBe(true);
  });
});

describe('getSubstrateTokenStatus', () => {
  it('reports no token when session has none', () => {
    vi.mocked(readSessionState).mockReturnValue(null);
    expect(getSubstrateTokenStatus()).toEqual({ hasToken: false });
  });

  it('reports token details when present', () => {
    const token = makeJwt({ exp: futureExp() });
    vi.mocked(readSessionState).mockReturnValue(
      sessionWith([entry('k', { target: 'substrate.office.com/SubstrateSearch', secret: token })]),
    );
    const status = getSubstrateTokenStatus();
    expect(status.hasToken).toBe(true);
    expect(typeof status.expiresAt).toBe('string');
    expect(status.minutesRemaining).toBeGreaterThan(0);
  });
});

// ============================================================================
// extractTeamsToken
// ============================================================================

describe('extractTeamsToken', () => {
  it('prefers the chatsvc candidate and derives userMri from oid', () => {
    const chatToken = makeJwt({ exp: futureExp(), oid: 'oid-guid' });
    const state = sessionWith([
      entry('c', { target: 'https://chatsvcagg.teams.microsoft.com/.default', secret: chatToken }),
    ]);
    const result = extractTeamsToken(state);
    expect(result?.token).toBe(chatToken);
    expect(result?.userMri).toBe('8:orgid:oid-guid');
  });

  it('falls back to the skype candidate when no chatsvc token exists', () => {
    const skypeToken = makeJwt({ exp: futureExp(), oid: 'oid2' });
    const state = sessionWith([
      entry('s', { target: 'https://api.spaces.skype.com/.default', secret: skypeToken }),
    ]);
    expect(extractTeamsToken(state)?.token).toBe(skypeToken);
  });

  it('derives userMri from substrate token when chat tokens lack oid', () => {
    const chatToken = makeJwt({ exp: futureExp() }); // no oid
    const substrateToken = makeJwt({ exp: futureExp(), oid: 'sub-oid' });
    const state = sessionWith([
      entry('c', { target: 'https://chatsvcagg.teams.microsoft.com/.default', secret: chatToken }),
      entry('sub', { target: 'substrate.office.com/SubstrateSearch', secret: substrateToken }),
    ]);
    expect(extractTeamsToken(state)?.userMri).toBe('8:orgid:sub-oid');
  });

  it('keeps the chatsvc token with the latest expiry', () => {
    const older = makeJwt({ exp: futureExp(HOUR), oid: 'o' });
    const newer = makeJwt({ exp: futureExp(HOUR * 4), oid: 'o' });
    const state = sessionWith([
      entry('c1', { target: 'chatsvcagg.teams.microsoft.com', secret: older }),
      entry('c2', { target: 'chatsvcagg.teams.microsoft.com', secret: newer }),
    ]);
    expect(extractTeamsToken(state)?.token).toBe(newer);
  });

  it('returns null when no userMri can be resolved', () => {
    const chatToken = makeJwt({ exp: futureExp() }); // no oid, no substrate fallback
    const state = sessionWith([
      entry('c', { target: 'chatsvcagg.teams.microsoft.com', secret: chatToken }),
    ]);
    expect(extractTeamsToken(state)).toBeNull();
  });

  it('returns null when the best candidate is expired', () => {
    const chatToken = makeJwt({ exp: pastExp(), oid: 'o' });
    const state = sessionWith([
      entry('c', { target: 'chatsvcagg.teams.microsoft.com', secret: chatToken }),
    ]);
    expect(extractTeamsToken(state)).toBeNull();
  });

  it('skips entries without a target, non-jwt secret or no exp', () => {
    const state = sessionWith([
      entry('bad', 'not-json{'),
      entry('no-target', { secret: makeJwt({ exp: futureExp(), oid: 'o' }) }),
      entry('not-jwt', { target: 'chatsvcagg.teams.microsoft.com', secret: 'plain' }),
      entry('no-exp', { target: 'chatsvcagg.teams.microsoft.com', secret: makeJwt({ oid: 'o' }) }),
    ]);
    expect(extractTeamsToken(state)).toBeNull();
  });
});

// ============================================================================
// extractSkypeSpacesToken
// ============================================================================

describe('extractSkypeSpacesToken', () => {
  it('returns the latest-expiry skype spaces token', () => {
    const older = makeJwt({ exp: futureExp(HOUR) });
    const newer = makeJwt({ exp: futureExp(HOUR * 3) });
    const state = sessionWith([
      entry('a', { target: 'https://api.spaces.skype.com/.default', secret: older }),
      entry('b', { target: 'https://api.spaces.skype.com/Authorization.ReadWrite', secret: newer }),
    ]);
    expect(extractSkypeSpacesToken(state)).toBe(newer);
  });

  it('skips unrelated, expired, non-jwt and no-exp entries; returns null', () => {
    const state = sessionWith([
      entry('bad', 'not-json{'),
      entry('other', { target: 'https://graph.microsoft.com', secret: makeJwt({ exp: futureExp() }) }),
      entry('not-jwt', { target: 'api.spaces.skype.com', secret: 'plain' }),
      entry('no-exp', { target: 'api.spaces.skype.com', secret: makeJwt({ sub: 'x' }) }),
      entry('expired', { target: 'api.spaces.skype.com', secret: makeJwt({ exp: pastExp() }) }),
    ]);
    expect(extractSkypeSpacesToken(state)).toBeNull();
  });

  it('returns null when no session', () => {
    vi.mocked(readSessionState).mockReturnValue(null);
    expect(extractSkypeSpacesToken()).toBeNull();
  });
});

// ============================================================================
// extractGraphToken
// ============================================================================

describe('extractGraphToken', () => {
  it('returns a graph token that grants Calendars.ReadWrite', () => {
    const token = makeJwt({ exp: futureExp(), scp: 'Calendars.ReadWrite Mail.Read' });
    const state = sessionWith([
      entry('g', { target: 'https://graph.microsoft.com/.default', secret: token }),
    ]);
    expect(extractGraphToken(state)).toBe(token);
  });

  it('stays lenient when scopes cannot be determined', () => {
    const token = makeJwt({ exp: futureExp() }); // no scp/roles -> null grant
    const state = sessionWith([
      entry('g', { target: 'graph.microsoft.com', secret: token }),
    ]);
    expect(extractGraphToken(state)).toBe(token);
  });

  it('skips graph tokens that explicitly lack calendar write access', () => {
    const token = makeJwt({ exp: futureExp(), scp: 'Mail.Read User.Read' });
    const state = sessionWith([
      entry('g', { target: 'graph.microsoft.com', secret: token }),
    ]);
    expect(extractGraphToken(state)).toBeNull();
  });

  it('skips unrelated/expired/non-jwt entries and returns null', () => {
    const state = sessionWith([
      entry('bad', 'not-json{'),
      entry('other', { target: 'api.spaces.skype.com', secret: makeJwt({ exp: futureExp() }) }),
      entry('not-jwt', { target: 'graph.microsoft.com', secret: 'plain' }),
      entry('no-exp', { target: 'graph.microsoft.com', secret: makeJwt({ scp: 'Calendars.ReadWrite' }) }),
      entry('expired', { target: 'graph.microsoft.com', secret: makeJwt({ exp: pastExp(), scp: 'Calendars.ReadWrite' }) }),
    ]);
    expect(extractGraphToken(state)).toBeNull();
  });
});

// ============================================================================
// extractRegionConfig
// ============================================================================

describe('extractRegionConfig', () => {
  it('parses a partitioned middleTier config', () => {
    const state = sessionWith([
      entry('ts.DISCOVER-REGION-GTM', {
        item: {
          middleTier: 'https://teams.microsoft.com/api/mt/part/amer-02',
          chatServiceAfd: 'https://teams.microsoft.com/api/chatsvc/amer',
          chatSvcAggAfd: 'https://teams.microsoft.com/api/csa/amer',
        },
      }),
    ]);
    const cfg = extractRegionConfig(state);
    expect(cfg).toMatchObject({
      region: 'amer',
      partition: '02',
      regionPartition: 'amer-02',
      hasPartition: true,
      teamsBaseUrl: 'https://teams.microsoft.com',
    });
  });

  it('parses a non-partitioned middleTier config', () => {
    const state = sessionWith([
      entry('ts.DISCOVER-REGION-GTM', {
        item: {
          middleTier: 'https://teams.microsoft.us/api/mt/emea',
          chatServiceAfd: 'https://teams.microsoft.us/api/chatsvc/uk',
        },
      }),
    ]);
    const cfg = extractRegionConfig(state);
    expect(cfg?.region).toBe('uk');
    expect(cfg?.hasPartition).toBe(false);
    expect(cfg?.regionPartition).toBe('emea');
    expect(cfg?.teamsBaseUrl).toBe('https://teams.microsoft.us');
    expect(cfg?.csaServiceUrl).toBe('https://teams.microsoft.us/api/csa/uk');
  });

  it('handles missing middleTier (regionPartition falls back to region)', () => {
    const state = sessionWith([
      entry('ts.DISCOVER-REGION-GTM', {
        item: { chatServiceAfd: 'https://teams.microsoft.com/api/chatsvc/amer' },
      }),
    ]);
    const cfg = extractRegionConfig(state);
    expect(cfg?.regionPartition).toBe('amer');
    expect(cfg?.middleTierUrl).toBe('');
  });

  it('uses the fallback base URL when chatServiceAfd is not a valid URL', () => {
    const state = sessionWith([
      entry('ts.DISCOVER-REGION-GTM', {
        item: { chatServiceAfd: 'not a url /api/chatsvc/amer' },
      }),
    ]);
    const cfg = extractRegionConfig(state);
    expect(cfg?.teamsBaseUrl).toBe('https://teams.microsoft.com');
    expect(cfg?.region).toBe('amer');
  });

  it('skips when chatServiceAfd is missing or does not match the region pattern', () => {
    const noChat = sessionWith([entry('ts.DISCOVER-REGION-GTM', { item: { middleTier: 'x' } })]);
    expect(extractRegionConfig(noChat)).toBeNull();

    const badPattern = sessionWith([
      entry('ts.DISCOVER-REGION-GTM', { item: { chatServiceAfd: 'https://teams.microsoft.com/api/other/amer' } }),
    ]);
    expect(extractRegionConfig(badPattern)).toBeNull();
  });

  it('continues past unparseable DISCOVER-REGION-GTM entries and returns null', () => {
    const state = sessionWith([entry('ts.DISCOVER-REGION-GTM', 'not-json{')]);
    expect(extractRegionConfig(state)).toBeNull();
  });

  it('returns null when no region config key is present', () => {
    const state = sessionWith([entry('other', { foo: 'bar' })]);
    expect(extractRegionConfig(state)).toBeNull();
  });
});

// ============================================================================
// extractUserDetails
// ============================================================================

describe('extractUserDetails', () => {
  it('parses full user details with licenses', () => {
    const state = sessionWith([
      entry('ts.DISCOVER-USER-DETAILS', {
        item: {
          id: '8:orgid:abc',
          region: 'amer',
          userPartition: 'amer01',
          partition: 'amer02',
          licenseDetails: {
            isFreemium: false,
            isTrial: true,
            isTeamsEnabled: true,
            isCopilot: true,
            isTranscriptEnabled: false,
            isFrontline: false,
          },
        },
      }),
    ]);
    const details = extractUserDetails(state);
    expect(details?.mri).toBe('8:orgid:abc');
    expect(details?.region).toBe('amer');
    expect(details?.userPartition).toBe('amer01');
    expect(details?.tenantPartition).toBe('amer02');
    expect(details?.licenses.isCopilot).toBe(true);
    expect(details?.licenses.isTrial).toBe(true);
  });

  it('defaults licenses and partitions when fields are missing', () => {
    const state = sessionWith([
      entry('ts.DISCOVER-USER-DETAILS', { item: { id: '8:orgid:x', region: 'emea' } }),
    ]);
    const details = extractUserDetails(state);
    expect(details?.userPartition).toBe('');
    expect(details?.tenantPartition).toBe('');
    expect(details?.licenses.isFreemium).toBe(false);
  });

  it('skips entries missing id or region, and unparseable entries', () => {
    const missing = sessionWith([entry('ts.DISCOVER-USER-DETAILS', { item: { id: 'x' } })]);
    expect(extractUserDetails(missing)).toBeNull();

    const bad = sessionWith([entry('ts.DISCOVER-USER-DETAILS', 'not-json{')]);
    expect(extractUserDetails(bad)).toBeNull();
  });

  it('returns null when no user-details key is present', () => {
    expect(extractUserDetails(sessionWith([entry('other', { a: 1 })]))).toBeNull();
  });
});

// ============================================================================
// extractMessageAuth / getMessageAuthStatus
// ============================================================================

describe('extractMessageAuth', () => {
  it('returns null when no session state', () => {
    expect(extractMessageAuth(null as unknown as SessionState)).toBeNull();
  });

  it('extracts skype + auth tokens and derives MRI from skypeid (orgid form)', () => {
    const skypeToken = makeJwt({ skypeid: 'orgid:guid-1' });
    const state: SessionState = {
      origins: [],
      cookies: [
        { name: 'skypetoken_asm', value: skypeToken, domain: 'teams.microsoft.com' },
        { name: 'authtoken', value: 'authval', domain: 'teams.microsoft.com' },
      ],
    };
    const auth = extractMessageAuth(state);
    expect(auth?.skypeToken).toBe(skypeToken);
    expect(auth?.userMri).toBe('8:orgid:guid-1');
  });

  it('decodes Bearer= prefixed authtoken and uses full-form skypeid', () => {
    const skypeToken = makeJwt({ skypeid: '8:orgid:full-mri' });
    const authJwt = makeJwt({ oid: 'auth-oid' });
    const state: SessionState = {
      origins: [],
      cookies: [
        { name: 'skypetoken_asm', value: skypeToken, domain: '.teams.microsoft.com' },
        { name: 'authtoken', value: encodeURIComponent(`Bearer=${authJwt}`), domain: 'teams.microsoft.com' },
      ],
    };
    const auth = extractMessageAuth(state);
    expect(auth?.authToken).toBe(authJwt);
    expect(auth?.userMri).toBe('8:orgid:full-mri');
  });

  it('falls back to authtoken oid when skype token has no skypeid', () => {
    const skypeToken = makeJwt({ sub: 'no-skypeid' });
    const authJwt = makeJwt({ oid: 'auth-oid' });
    const state: SessionState = {
      origins: [],
      cookies: [
        { name: 'skypetoken_asm', value: skypeToken, domain: 'teams.microsoft.com' },
        { name: 'authtoken', value: authJwt, domain: 'teams.microsoft.com' },
      ],
    };
    expect(extractMessageAuth(state)?.userMri).toBe('8:orgid:auth-oid');
  });

  it('keeps a plain skypeid that is neither orgid nor 8: prefixed', () => {
    const skypeToken = makeJwt({ skypeid: 'live:somebody' });
    const state: SessionState = {
      origins: [],
      cookies: [
        { name: 'skypetoken_asm', value: skypeToken, domain: 'teams.microsoft.com' },
        { name: 'authtoken', value: 'x', domain: 'teams.microsoft.com' },
      ],
    };
    expect(extractMessageAuth(state)?.userMri).toBe('live:somebody');
  });

  it('returns null when required cookies are missing', () => {
    const state: SessionState = {
      origins: [],
      cookies: [{ name: 'skypetoken_asm', value: 'x', domain: 'teams.microsoft.com' }],
    };
    expect(extractMessageAuth(state)).toBeNull();
  });

  it('returns null when no userMri can be derived', () => {
    const skypeToken = makeJwt({ sub: 'nope' });
    const authToken = makeJwt({ sub: 'nope' });
    const state: SessionState = {
      origins: [],
      cookies: [
        { name: 'skypetoken_asm', value: skypeToken, domain: 'teams.microsoft.com' },
        { name: 'authtoken', value: authToken, domain: 'teams.microsoft.com' },
      ],
    };
    expect(extractMessageAuth(state)).toBeNull();
  });
});

describe('getMessageAuthStatus', () => {
  it('reports no token when session missing', () => {
    vi.mocked(readSessionState).mockReturnValue(null);
    expect(getMessageAuthStatus()).toEqual({ hasToken: false });
  });

  it('reports no token when skype cookie absent', () => {
    vi.mocked(readSessionState).mockReturnValue({ origins: [], cookies: [] });
    expect(getMessageAuthStatus()).toEqual({ hasToken: false });
  });

  it('reports valid when cookie present but expiry unparseable', () => {
    vi.mocked(readSessionState).mockReturnValue({
      origins: [],
      cookies: [{ name: 'skypetoken_asm', value: 'opaque-token', domain: 'teams.microsoft.com' }],
    });
    expect(getMessageAuthStatus()).toEqual({ hasToken: true });
  });

  it('reports expiry details when cookie is a JWT', () => {
    const token = makeJwt({ exp: futureExp() });
    vi.mocked(readSessionState).mockReturnValue({
      origins: [],
      cookies: [{ name: 'skypetoken_asm', value: token, domain: 'teams.microsoft.com' }],
    });
    const status = getMessageAuthStatus();
    expect(status.hasToken).toBe(true);
    expect(status.minutesRemaining).toBeGreaterThan(0);
  });
});

// ============================================================================
// extractCsaToken
// ============================================================================

describe('extractCsaToken', () => {
  it('returns null when no session state', () => {
    vi.mocked(readSessionState).mockReturnValue(null);
    expect(extractCsaToken()).toBeNull();
  });

  it('finds a chatsvcagg secret across origins, skipping tmp. and non-matching keys', () => {
    const state: SessionState = {
      cookies: [],
      origins: [
        { origin: 'https://other.com', localStorage: [entry('tmp.chatsvcagg.teams.microsoft.com', { secret: 'ignored' })] },
        {
          origin: 'https://teams.microsoft.com',
          localStorage: [
            entry('unrelated', { secret: 'no' }),
            entry('bad', 'not-json{'),
            entry('x-chatsvcagg.teams.microsoft.com-y', { secret: 'the-csa-token' }),
          ],
        },
      ],
    };
    expect(extractCsaToken(state)).toBe('the-csa-token');
  });

  it('returns null when matching entry has no secret', () => {
    const state: SessionState = {
      cookies: [],
      origins: [{ origin: 'x', localStorage: [entry('chatsvcagg.teams.microsoft.com', { notsecret: 'y' })] }],
    };
    expect(extractCsaToken(state)).toBeNull();
  });
});

// ============================================================================
// getUserProfile / getUserDisplayName
// ============================================================================

describe('getUserProfile', () => {
  it('returns a profile parsed from a JWT secret', () => {
    const token = makeJwt({ oid: 'pid', name: 'Doe, John', upn: 'john@x.com', tid: 'tenant' });
    const state = sessionWith([
      entry('bad', 'not-json{'),
      entry('not-jwt', { secret: 'plain' }),
      entry('jwt', { secret: token }),
    ]);
    const profile = getUserProfile(state);
    expect(profile?.id).toBe('pid');
    expect(profile?.email).toBe('john@x.com');
  });

  it('returns null when no JWT yields a profile', () => {
    const state = sessionWith([entry('jwt', { secret: makeJwt({ foo: 'bar' }) })]);
    expect(getUserProfile(state)).toBeNull();
  });
});

describe('getUserDisplayName', () => {
  it('returns explicit displayName from localStorage', () => {
    const state = sessionWith([entry('k', { displayName: 'Jane Doe' })]);
    expect(getUserDisplayName(state)).toBe('Jane Doe');
  });

  it('returns nested name.displayName from localStorage', () => {
    const state = sessionWith([entry('k', { givenName: 'x', name: { displayName: 'Nested Name' } })]);
    expect(getUserDisplayName(state)).toBe('Nested Name');
  });

  it('falls back to the Teams token name claim', () => {
    const chatToken = makeJwt({ exp: futureExp(), oid: 'o', name: 'Token Name' });
    const state = sessionWith([
      entry('c', { target: 'chatsvcagg.teams.microsoft.com', secret: chatToken }),
    ]);
    expect(getUserDisplayName(state)).toBe('Token Name');
  });

  it('returns null when nothing yields a name', () => {
    const state = sessionWith([entry('k', { other: 'value' })]);
    expect(getUserDisplayName(state)).toBeNull();
  });

  it('skips unparseable displayName-like entries', () => {
    const state = sessionWith([entry('k', 'displayName but not json {')]);
    expect(getUserDisplayName(state)).toBeNull();
  });
});

// ============================================================================
// areTokensExpired
// ============================================================================

describe('areTokensExpired', () => {
  it('returns true when no session state', () => {
    expect(areTokensExpired(null as unknown as SessionState)).toBe(true);
    vi.mocked(readSessionState).mockReturnValue(null);
    expect(areTokensExpired()).toBe(true);
  });

  it('returns false when a valid substrate token exists', () => {
    const token = makeJwt({ exp: futureExp() });
    const state = sessionWith([entry('k', { target: 'substrate.office.com/SubstrateSearch', secret: token })]);
    expect(areTokensExpired(state)).toBe(false);
  });

  it('returns true when substrate token cannot be found', () => {
    const state = sessionWith([entry('k', { target: 'graph.microsoft.com', secret: makeJwt({ exp: futureExp() }) })]);
    expect(areTokensExpired(state)).toBe(true);
  });
});

// ============================================================================
// discoverConfig
// ============================================================================

describe('discoverConfig', () => {
  it('returns null when no session/localStorage', () => {
    vi.mocked(readSessionState).mockReturnValue(null);
    expect(discoverConfig()).toBeNull();
  });

  it('collects discovery configs, config keys, urls, hosts and base/substrate URLs', () => {
    const state = sessionWith([
      entry('ts.DISCOVER-REGION-GTM', {
        item: { chatServiceAfd: 'https://teams.microsoft.com/api/chatsvc/amer' },
      }),
      entry('app.DISCOVER-BAD', 'not-json-value'),
      entry('settings.flags', { url: 'https://substrate.office.com/searchservice' }),
      entry('short-config', 'tiny'),
      entry('endpoint.accesstoken', { nested: { link: 'https://chatsvcagg.teams.microsoft.com/x' } }),
      entry('plain', 'no config here'),
    ]);
    const cfg = discoverConfig(state);
    expect(cfg).not.toBeNull();
    expect(Object.keys(cfg!.discoveryConfigs)).toContain('ts.DISCOVER-REGION-GTM');
    // Non-JSON DISCOVER value stored as raw string
    expect(cfg!.discoveryConfigs['app.DISCOVER-BAD']).toBe('not-json-value');
    expect(cfg!.configKeys).toContain('settings.flags');
    expect(cfg!.teamsBaseUrl).toBe('https://teams.microsoft.com');
    expect(cfg!.substrateUrl).toBe('https://substrate.office.com');
    expect(cfg!.uniqueHosts).toContain('substrate.office.com');
    // configContents captured the short non-JSON config value
    expect(cfg!.configContents['short-config']).toBe('tiny');
  });
});
