/**
 * Unit tests for people tools (real handlers).
 */
import { describe, it, expect, vi, beforeEach } from 'vitest';
import { ok, err } from '../types/result.js';
import { createError, ErrorCode } from '../types/errors.js';
import './index.js';

vi.mock('../api/substrate-api.js', () => ({
  searchPeople: vi.fn(),
  getFrequentContacts: vi.fn(),
}));
vi.mock('../api/presence-api.js', () => ({ getPresence: vi.fn() }));
vi.mock('../auth/token-extractor.js', () => ({ getUserProfile: vi.fn() }));

import { searchPeople, getFrequentContacts } from '../api/substrate-api.js';
import { getPresence } from '../api/presence-api.js';
import { getUserProfile } from '../auth/token-extractor.js';
import {
  getMeTool,
  searchPeopleTool,
  frequentContactsTool,
  getPresenceTool,
  SearchPeopleInputSchema,
  FrequentContactsInputSchema,
  GetPresenceInputSchema,
} from './people-tools.js';

const ctx = { server: {} } as never;
const anErr = err(createError(ErrorCode.API_ERROR, 'boom'));

beforeEach(() => vi.clearAllMocks());

describe('getMeTool', () => {
  it('returns the profile when present', async () => {
    vi.mocked(getUserProfile).mockReturnValue({ name: 'Shayan' } as never);
    const res = await getMeTool.handler({}, ctx);
    expect(res.success).toBe(true);
    if (res.success) expect(res.data).toEqual({ profile: { name: 'Shayan' } });
  });

  it('returns AUTH_REQUIRED when no profile', async () => {
    vi.mocked(getUserProfile).mockReturnValue(null as never);
    const res = await getMeTool.handler({}, ctx);
    expect(res.success).toBe(false);
    if (!res.success) expect(res.error.code).toBe(ErrorCode.AUTH_REQUIRED);
  });
});

describe('searchPeopleTool', () => {
  it('parses defaults', () => {
    const parsed = SearchPeopleInputSchema.parse({ query: 'rob' });
    expect(parsed.limit).toBe(10);
  });

  it('returns results on success', async () => {
    vi.mocked(searchPeople).mockResolvedValue(ok({ returned: 1, results: [{ id: 'p1' }] }) as never);
    const res = await searchPeopleTool.handler({ query: 'rob', limit: 5 }, ctx);
    expect(res.success).toBe(true);
    if (res.success) expect(res.data).toEqual({ query: 'rob', returned: 1, results: [{ id: 'p1' }] });
  });

  it('propagates errors', async () => {
    vi.mocked(searchPeople).mockResolvedValue(anErr as never);
    const res = await searchPeopleTool.handler({ query: 'x', limit: 10 }, ctx);
    expect(res.success).toBe(false);
  });
});

describe('frequentContactsTool', () => {
  it('parses defaults', () => {
    expect(FrequentContactsInputSchema.parse({}).limit).toBe(50);
  });

  it('returns contacts on success', async () => {
    vi.mocked(getFrequentContacts).mockResolvedValue(ok({ returned: 2, results: [{ id: 'c1' }] }) as never);
    const res = await frequentContactsTool.handler({ limit: 50 }, ctx);
    expect(res.success).toBe(true);
    if (res.success) expect(res.data).toEqual({ returned: 2, contacts: [{ id: 'c1' }] });
  });

  it('propagates errors', async () => {
    vi.mocked(getFrequentContacts).mockResolvedValue(anErr as never);
    const res = await frequentContactsTool.handler({ limit: 50 }, ctx);
    expect(res.success).toBe(false);
  });
});

describe('getPresenceTool', () => {
  it('validates userIds', () => {
    expect(() => GetPresenceInputSchema.parse({ userIds: [] })).toThrow();
  });

  it('returns presences on success', async () => {
    vi.mocked(getPresence).mockResolvedValue(ok([{ id: 'u1', availability: 'Available' }]) as never);
    const res = await getPresenceTool.handler({ userIds: ['u1'] }, ctx);
    expect(res.success).toBe(true);
    if (res.success) expect(res.data).toEqual({ presences: [{ id: 'u1', availability: 'Available' }] });
  });

  it('propagates errors', async () => {
    vi.mocked(getPresence).mockResolvedValue(anErr as never);
    const res = await getPresenceTool.handler({ userIds: ['u1'] }, ctx);
    expect(res.success).toBe(false);
  });
});
