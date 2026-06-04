/**
 * Unit tests for tools/index handleApiResult helper.
 */
import { describe, it, expect } from 'vitest';
import { ok, err } from '../types/result.js';
import { createError, ErrorCode } from '../types/errors.js';
import { handleApiResult } from './index.js';

describe('handleApiResult', () => {
  it('transforms an ok result into a success ToolResult', () => {
    const result = handleApiResult(ok({ count: 2 }), (v) => ({ total: v.count }));
    expect(result.success).toBe(true);
    if (result.success) expect(result.data).toEqual({ total: 2 });
  });

  it('passes the error through for an err result', () => {
    const error = createError(ErrorCode.NOT_FOUND, 'missing');
    const result = handleApiResult(err(error), () => ({ never: true }));
    expect(result.success).toBe(false);
    if (!result.success) {
      expect(result.error).toBe(error);
      expect(result.error.code).toBe(ErrorCode.NOT_FOUND);
    }
  });
});
