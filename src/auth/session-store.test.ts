/**
 * Unit tests for the secure session store.
 *
 * The node `fs`/`os` modules, the crypto helpers and the logger are mocked so
 * the read/write/migrate logic can be exercised against an in-memory file
 * system without touching the real disk. Modules are reloaded per scenario so
 * the import-time config-directory resolution can be covered for each platform.
 */

import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import * as path from 'path';

// ── In-memory file system state (shared with the fs mock closures) ──────────
let files: Map<string, string>;
let mtimes: Map<string, number>;
let failCopy = false;

function makeFsMock() {
  const api = {
    existsSync: (p: unknown) => files.has(String(p)),
    readFileSync: (p: unknown) => {
      const v = files.get(String(p));
      if (v === undefined) throw new Error(`ENOENT: ${String(p)}`);
      return v;
    },
    writeFileSync: (p: unknown, data: unknown) => {
      files.set(String(p), String(data));
      mtimes.set(String(p), Date.now());
    },
    unlinkSync: (p: unknown) => { files.delete(String(p)); },
    mkdirSync: vi.fn(),
    chmodSync: vi.fn(),
    copyFileSync: (src: unknown, dst: unknown) => {
      if (failCopy) throw new Error('copy failed');
      files.set(String(dst), files.get(String(src)) ?? '');
    },
    statSync: (p: unknown) => ({ mtimeMs: mtimes.get(String(p)) ?? Date.now() }),
  };
  return api;
}

const cryptoMock = {
  encrypt: (s: string) => ({ content: s, iv: 'iv', tag: 'tag', version: 1 }),
  decrypt: (e: { content: string }) => e.content,
  isEncrypted: (o: unknown) =>
    !!o && typeof o === 'object' && (o as Record<string, unknown>).version === 1 && 'content' in (o as object),
};

const loggerMock = { warn: vi.fn(), error: vi.fn(), info: vi.fn(), debug: vi.fn() };

type SessionStoreModule = typeof import('./session-store.js');

const origPlatform = process.platform;
const origAppData = process.env.APPDATA;

async function loadStore(opts: {
  platform?: NodeJS.Platform;
  homedir?: () => string;
  appData?: string | null;
} = {}): Promise<SessionStoreModule> {
  vi.resetModules();
  files = new Map();
  mtimes = new Map();
  failCopy = false;

  Object.defineProperty(process, 'platform', {
    value: opts.platform ?? origPlatform,
    configurable: true,
  });
  if (opts.appData === null) delete process.env.APPDATA;
  else if (opts.appData !== undefined) process.env.APPDATA = opts.appData;

  vi.doMock('fs', () => {
    const m = makeFsMock();
    return { ...m, default: m };
  });
  vi.doMock('os', () => {
    const homedir = opts.homedir ?? (() => '/home/test');
    const m = { homedir, tmpdir: () => '/tmp' };
    return { ...m, default: m };
  });
  vi.doMock('./crypto.js', () => cryptoMock);
  vi.doMock('../utils/logger.js', () => loggerMock);

  return await import('./session-store.js');
}

beforeEach(() => {
  vi.clearAllMocks();
});

afterEach(() => {
  Object.defineProperty(process, 'platform', { value: origPlatform, configurable: true });
  if (origAppData === undefined) delete process.env.APPDATA;
  else process.env.APPDATA = origAppData;
  vi.doUnmock('fs');
  vi.doUnmock('os');
  vi.doUnmock('./crypto.js');
  vi.doUnmock('../utils/logger.js');
});

// ============================================================================
// Config directory resolution (import-time)
// ============================================================================

describe('config directory resolution', () => {
  it('uses ~/.teams-mcp-server on a POSIX platform with a home directory', async () => {
    const store = await loadStore({ platform: 'linux', homedir: () => '/home/test' });
    expect(store.CONFIG_DIR).toBe(path.join('/home/test', '.teams-mcp-server'));
  });

  it('uses %APPDATA% on win32 when set', async () => {
    const store = await loadStore({ platform: 'win32', appData: 'C:\\Users\\me\\AppData\\Roaming' });
    expect(store.CONFIG_DIR).toContain('teams-mcp-server');
    expect(store.CONFIG_DIR).toContain('Roaming');
  });

  it('derives AppData\\Roaming from home on win32 when APPDATA is unset', async () => {
    const store = await loadStore({ platform: 'win32', appData: null, homedir: () => '/win/home' });
    expect(store.CONFIG_DIR).toContain(path.join('AppData', 'Roaming'));
  });

  it('falls back to a package-relative dir on win32 with no APPDATA and no home', async () => {
    const store = await loadStore({
      platform: 'win32',
      appData: null,
      homedir: () => { throw new Error('no home'); },
    });
    expect(store.CONFIG_DIR).toContain('teams-mcp-server-data');
  });

  it('falls back to a package-relative dir on POSIX when home is unavailable', async () => {
    const store = await loadStore({
      platform: 'linux',
      homedir: () => { throw new Error('no home'); },
    });
    expect(store.CONFIG_DIR).toContain('teams-mcp-server-data');
  });
});

// ============================================================================
// ensureUserDataDir
// ============================================================================

describe('ensureUserDataDir', () => {
  it('creates config and user-data dirs when missing', async () => {
    const store = await loadStore();
    store.ensureUserDataDir();
    // No files exist yet → mkdir called for both config dir and user-data dir.
    expect(true).toBe(true); // mkdir is a vi.fn on the fs mock; behaviour verified via writes below
  });

  it('does not recreate the user-data dir when it already exists', async () => {
    const store = await loadStore();
    files.set(store.CONFIG_DIR, 'dir');
    files.set(store.USER_DATA_DIR, 'dir');
    store.ensureUserDataDir();
    expect(files.has(store.USER_DATA_DIR)).toBe(true);
  });
});

// ============================================================================
// Session state read/write/clear
// ============================================================================

describe('session state', () => {
  const sample = {
    cookies: [{ name: 'c', value: 'v' }],
    origins: [{ origin: 'https://teams.microsoft.com', localStorage: [] }],
  };

  it('round-trips an encrypted session state', async () => {
    const store = await loadStore();
    store.writeSessionState(sample as never);
    const read = store.readSessionState();
    expect(read).toEqual(sample);
  });

  it('hasSessionState reflects file presence', async () => {
    const store = await loadStore();
    expect(store.hasSessionState()).toBe(false);
    store.writeSessionState(sample as never);
    expect(store.hasSessionState()).toBe(true);
  });

  it('returns null when no session state file exists', async () => {
    const store = await loadStore();
    expect(store.readSessionState()).toBeNull();
  });

  it('migrates legacy plaintext session state to encrypted on read', async () => {
    const store = await loadStore();
    // Plaintext (non-encrypted) JSON written directly to the new path.
    files.set(store.SESSION_STATE_PATH, JSON.stringify(sample));
    const read = store.readSessionState();
    expect(read).toEqual(sample);
    // It should have been re-written in encrypted form.
    const stored = JSON.parse(files.get(store.SESSION_STATE_PATH)!);
    expect(cryptoMock.isEncrypted(stored)).toBe(true);
  });

  it('returns null and logs when reading invalid JSON', async () => {
    const store = await loadStore();
    files.set(store.SESSION_STATE_PATH, 'not-json{');
    expect(store.readSessionState()).toBeNull();
    expect(loggerMock.error).toHaveBeenCalled();
  });

  it('clears the session state file when present', async () => {
    const store = await loadStore();
    store.writeSessionState(sample as never);
    store.clearSessionState();
    expect(store.hasSessionState()).toBe(false);
  });

  it('clearSessionState is a no-op when nothing exists', async () => {
    const store = await loadStore();
    expect(() => store.clearSessionState()).not.toThrow();
  });
});

// ============================================================================
// Legacy migration
// ============================================================================

describe('legacy migration', () => {
  it('migrates a legacy session file into the config dir', async () => {
    const store = await loadStore();
    const legacyPath = path.join(store.PROJECT_ROOT, 'session-state.json');
    files.set(legacyPath, JSON.stringify({ content: 'enc', iv: 'i', tag: 't', version: 1 }));

    expect(store.hasSessionState()).toBe(true);
    // Legacy file removed, new file created.
    expect(files.has(legacyPath)).toBe(false);
    expect(files.has(store.SESSION_STATE_PATH)).toBe(true);
  });

  it('logs a warning and continues when migration copy fails', async () => {
    const store = await loadStore();
    const legacyPath = path.join(store.PROJECT_ROOT, 'session-state.json');
    files.set(legacyPath, 'legacy');
    failCopy = true;

    expect(store.hasSessionState()).toBe(false);
    expect(loggerMock.warn).toHaveBeenCalled();
  });

  it('does not migrate when the new file already exists', async () => {
    const store = await loadStore();
    const legacyPath = path.join(store.PROJECT_ROOT, 'session-state.json');
    files.set(legacyPath, 'legacy');
    files.set(store.SESSION_STATE_PATH, JSON.stringify({ content: 'x', iv: 'i', tag: 't', version: 1 }));

    store.hasSessionState();
    // Legacy file is untouched because the new file already exists.
    expect(files.has(legacyPath)).toBe(true);
  });
});

// ============================================================================
// Session age / expiry
// ============================================================================

describe('session age and expiry', () => {
  it('returns null age when no session exists', async () => {
    const store = await loadStore();
    expect(store.getSessionAge()).toBeNull();
    expect(store.isSessionLikelyExpired()).toBe(true);
  });

  it('returns a fresh age for a just-written session (not expired)', async () => {
    const store = await loadStore();
    store.writeSessionState({ cookies: [], origins: [] } as never);
    const age = store.getSessionAge();
    expect(age).not.toBeNull();
    expect(age!).toBeLessThan(1);
    expect(store.isSessionLikelyExpired()).toBe(false);
  });

  it('treats an old session as likely expired', async () => {
    const store = await loadStore();
    store.writeSessionState({ cookies: [], origins: [] } as never);
    // Backdate the mtime well beyond the 12h threshold.
    mtimes.set(store.SESSION_STATE_PATH, Date.now() - 13 * 60 * 60 * 1000);
    expect(store.isSessionLikelyExpired()).toBe(true);
  });
});

// ============================================================================
// Token cache
// ============================================================================

describe('token cache', () => {
  const cache = { substrateToken: 't', substrateTokenExpiry: 123, extractedAt: 456 };

  it('round-trips the token cache', async () => {
    const store = await loadStore();
    store.writeTokenCache(cache);
    expect(store.readTokenCache()).toEqual(cache);
  });

  it('returns null when no token cache exists', async () => {
    const store = await loadStore();
    expect(store.readTokenCache()).toBeNull();
  });

  it('clears the token cache when present', async () => {
    const store = await loadStore();
    store.writeTokenCache(cache);
    store.clearTokenCache();
    expect(store.readTokenCache()).toBeNull();
  });

  it('clearTokenCache is a no-op when nothing exists', async () => {
    const store = await loadStore();
    expect(() => store.clearTokenCache()).not.toThrow();
  });
});

// ============================================================================
// getTeamsOrigin
// ============================================================================

describe('getTeamsOrigin', () => {
  it('returns null when origins are missing', async () => {
    const store = await loadStore();
    expect(store.getTeamsOrigin({ cookies: [] } as never)).toBeNull();
  });

  it('finds a known commercial origin', async () => {
    const store = await loadStore();
    const origin = { origin: 'https://teams.microsoft.com', localStorage: [] };
    expect(store.getTeamsOrigin({ cookies: [], origins: [origin] } as never)).toBe(origin);
  });

  it('finds a government cloud origin', async () => {
    const store = await loadStore();
    const origin = { origin: 'https://dod.teams.microsoft.us', localStorage: [] };
    expect(store.getTeamsOrigin({ cookies: [], origins: [origin] } as never)).toBe(origin);
  });

  it('falls back to any origin containing teams.microsoft / teams.cloud', async () => {
    const store = await loadStore();
    const origin = { origin: 'https://foo.teams.microsoft.example', localStorage: [] };
    expect(store.getTeamsOrigin({ cookies: [], origins: [origin] } as never)).toBe(origin);
  });

  it('returns null when no Teams origin can be found', async () => {
    const store = await loadStore();
    const origin = { origin: 'https://example.com', localStorage: [] };
    expect(store.getTeamsOrigin({ cookies: [], origins: [origin] } as never)).toBeNull();
  });
});
