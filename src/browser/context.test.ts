import { beforeEach, describe, expect, it, vi } from 'vitest';

const fsMocks = vi.hoisted(() => ({
  existsSync: vi.fn(),
  lstatSync: vi.fn(),
  readlinkSync: vi.fn(),
  unlinkSync: vi.fn(),
}));
const launchPersistentContext = vi.hoisted(() => vi.fn());

vi.mock('fs', () => fsMocks);
vi.mock('playwright', () => ({ chromium: { launchPersistentContext } }));
vi.mock('../auth/session-store.js', () => ({
  CONFIG_DIR: '/tmp/teams-mcp-test',
  ensureUserDataDir: vi.fn(),
  writeSessionState: vi.fn(),
}));
vi.mock('../utils/auth-guards.js', () => ({ clearRegionCache: vi.fn() }));
vi.mock('../utils/logger.js', () => ({ info: vi.fn(), warn: vi.fn(), error: vi.fn() }));

import { createBrowserContext } from './context.js';

function browserContext() {
  return {
    pages: () => [{ url: () => 'about:blank' }],
    newPage: vi.fn(),
  };
}

describe('createBrowserContext', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    fsMocks.lstatSync.mockReturnValue({} as never);
    fsMocks.readlinkSync.mockReturnValue('host-999999');
    vi.spyOn(process, 'kill').mockImplementation(() => { throw new Error('not running'); });
    launchPersistentContext.mockResolvedValue(browserContext());
  });

  it('removes a dangling Chrome profile lock before launching', async () => {
    await expect(createBrowserContext()).resolves.toBeDefined();
    expect(fsMocks.unlinkSync).toHaveBeenCalledWith('/tmp/teams-mcp-test/browser-profile/SingletonLock');
  });

  it('preserves a lock owned by a running browser', async () => {
    vi.mocked(process.kill).mockReturnValue(true);
    launchPersistentContext.mockRejectedValue(new Error('ProcessSingleton SingletonLock'));
    await expect(createBrowserContext()).rejects.toThrow('another MCP process is using its profile');
    expect(fsMocks.unlinkSync).not.toHaveBeenCalled();
  });
});
