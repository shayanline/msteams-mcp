/**
 * Unit tests for task tools (real handlers).
 */
import { describe, it, expect, vi, beforeEach } from 'vitest';
import { ok, err } from '../types/result.js';
import { createError, ErrorCode } from '../types/errors.js';
import './index.js';

vi.mock('../api/tasks-api.js', () => ({
  listTaskLists: vi.fn(),
  listTasks: vi.fn(),
  createTask: vi.fn(),
  setTaskCompleted: vi.fn(),
}));

import { listTaskLists, listTasks, createTask, setTaskCompleted } from '../api/tasks-api.js';
import {
  listTaskListsTool,
  listTasksTool,
  createTaskTool,
  completeTaskTool,
  ListTasksInputSchema,
} from './task-tools.js';

const ctx = { server: {} } as never;
const anErr = err(createError(ErrorCode.API_ERROR, 'boom'));

beforeEach(() => vi.clearAllMocks());

describe('listTaskListsTool', () => {
  it('returns lists on success', async () => {
    vi.mocked(listTaskLists).mockResolvedValue(ok({ lists: [{ id: 'l1' }] }) as never);
    const res = await listTaskListsTool.handler({}, ctx);
    expect(res.success).toBe(true);
    if (res.success) expect(res.data).toEqual({ lists: [{ id: 'l1' }] });
  });
  it('propagates errors', async () => {
    vi.mocked(listTaskLists).mockResolvedValue(anErr as never);
    expect((await listTaskListsTool.handler({}, ctx)).success).toBe(false);
  });
});

describe('listTasksTool', () => {
  it('parses defaults', () => {
    const parsed = ListTasksInputSchema.parse({});
    expect(parsed.includeCompleted).toBe(false);
    expect(parsed.top).toBe(50);
  });
  it('returns tasks on success', async () => {
    vi.mocked(listTasks).mockResolvedValue(ok({ listId: 'l1', tasks: [{ id: 't1' }] }) as never);
    const res = await listTasksTool.handler({ includeCompleted: false, top: 50 }, ctx);
    expect(res.success).toBe(true);
    if (res.success) expect(res.data).toEqual({ listId: 'l1', count: 1, tasks: [{ id: 't1' }] });
  });
  it('propagates errors', async () => {
    vi.mocked(listTasks).mockResolvedValue(anErr as never);
    expect((await listTasksTool.handler({ includeCompleted: false, top: 50 }, ctx)).success).toBe(false);
  });
});

describe('createTaskTool', () => {
  it('creates a task on success', async () => {
    vi.mocked(createTask).mockResolvedValue(ok({ id: 't1', title: 'x' }) as never);
    const res = await createTaskTool.handler({ title: 'x' }, ctx);
    expect(res.success).toBe(true);
    if (res.success) expect(res.data.message).toBe('Task created.');
  });
  it('propagates errors', async () => {
    vi.mocked(createTask).mockResolvedValue(anErr as never);
    expect((await createTaskTool.handler({ title: 'x' }, ctx)).success).toBe(false);
  });
});

describe('completeTaskTool', () => {
  it('completes a task', async () => {
    vi.mocked(setTaskCompleted).mockResolvedValue(ok({ id: 't1' }) as never);
    const res = await completeTaskTool.handler({ listId: 'l1', taskId: 't1', completed: true }, ctx);
    expect(res.success).toBe(true);
    if (res.success) expect(res.data.message).toBe('Task completed.');
  });
  it('reopens a task', async () => {
    vi.mocked(setTaskCompleted).mockResolvedValue(ok({ id: 't1' }) as never);
    const res = await completeTaskTool.handler({ listId: 'l1', taskId: 't1', completed: false }, ctx);
    expect(res.success).toBe(true);
    if (res.success) expect(res.data.message).toBe('Task reopened.');
  });
  it('propagates errors', async () => {
    vi.mocked(setTaskCompleted).mockResolvedValue(anErr as never);
    const res = await completeTaskTool.handler({ listId: 'l1', taskId: 't1', completed: true }, ctx);
    expect(res.success).toBe(false);
  });
});
