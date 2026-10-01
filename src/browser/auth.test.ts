import { describe, expect, it, vi } from 'vitest';
import { getAuthStatus } from './auth.js';

describe('getAuthStatus', () => {
  it('recognises the current Teams web app', async () => {
    const page = {
      url: () => 'https://teams.cloud.microsoft/v2/',
      locator: vi.fn(() => ({ count: vi.fn().mockResolvedValue(1) })),
    } as never;

    await expect(getAuthStatus(page)).resolves.toEqual({
      isAuthenticated: true,
      isOnLoginPage: false,
      currentUrl: 'https://teams.cloud.microsoft/v2/',
    });
  });
});
