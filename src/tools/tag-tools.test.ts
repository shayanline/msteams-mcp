/**
 * Unit tests for tag tools (real handlers).
 */
import { describe, it, expect, vi, beforeEach } from 'vitest';
import { ok, err } from '../types/result.js';
import { createError, ErrorCode } from '../types/errors.js';
import './index.js';

vi.mock('../api/tags-api.js', () => ({ listTeamTags: vi.fn() }));

import { listTeamTags } from '../api/tags-api.js';
import { tagTools } from './tag-tools.js';

const getTagsTool = tagTools[0];
const ctx = { server: {} } as never;

const tagsValue = {
  tags: [
    { id: 't1', displayName: 'Engineering', memberCount: 3, tagType: 'standard' },
    { id: 't2', displayName: 'Sales', memberCount: 5, tagType: 'standard' },
  ],
};

beforeEach(() => vi.clearAllMocks());

describe('getTagsTool', () => {
  it('returns all tags with mention syntax when no query', async () => {
    vi.mocked(listTeamTags).mockResolvedValue(ok(tagsValue) as never);
    const res = await getTagsTool.handler({ teamId: 'team1' }, ctx);
    expect(res.success).toBe(true);
    if (res.success) {
      expect(res.data.count).toBe(2);
      const tags = res.data.tags as Array<Record<string, unknown>>;
      expect(tags[0].mentionSyntax).toBe('@[Engineering](tag:t1)');
      expect(res.data.query).toBeUndefined();
    }
  });

  it('filters tags by query (case-insensitive)', async () => {
    vi.mocked(listTeamTags).mockResolvedValue(ok(tagsValue) as never);
    const res = await getTagsTool.handler({ teamId: 'team1', query: 'eng' }, ctx);
    expect(res.success).toBe(true);
    if (res.success) {
      expect(res.data.count).toBe(1);
      expect(res.data.query).toBe('eng');
    }
  });

  it('propagates errors', async () => {
    vi.mocked(listTeamTags).mockResolvedValue(err(createError(ErrorCode.API_ERROR, 'boom')) as never);
    expect((await getTagsTool.handler({ teamId: 'team1' }, ctx)).success).toBe(false);
  });
});
