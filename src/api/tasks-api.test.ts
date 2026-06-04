import { describe, it, expect, vi, beforeEach } from 'vitest';
import { ok } from '../types/result.js';

vi.mock('../utils/http.js', () => ({ httpRequest: vi.fn() }));
vi.mock('../utils/auth-guards.js', () => ({ requireGraphAuth: vi.fn(() => ok('graph-token')) }));

import { httpRequest } from '../utils/http.js';
import { requireGraphAuth } from '../utils/auth-guards.js';
import { listTaskLists, listTasks, createTask, setTaskCompleted } from './tasks-api.js';

const mockHttp = vi.mocked(httpRequest);
const mockAuth = vi.mocked(requireGraphAuth);
const httpOk = (data: unknown) => ok({ status: 200, headers: new Headers(), data } as never);

beforeEach(() => {
  vi.clearAllMocks();
  mockAuth.mockReturnValue(ok('graph-token') as never);
});

describe('listTaskLists', () => {
  it('maps lists and flags the default list', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ value: [
      { id: 'l1', displayName: 'Tasks', wellknownListName: 'defaultList' },
      { id: 'l2', displayName: 'Work' },
    ] }));
    const res = await listTaskLists();
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.lists).toEqual([
      { id: 'l1', displayName: 'Tasks', isDefault: true },
      { id: 'l2', displayName: 'Work', isDefault: false },
    ]);
    expect(mockHttp).toHaveBeenCalledWith(expect.stringContaining('/me/todo/lists'), expect.objectContaining({ method: 'GET' }));
  });

  it('propagates auth failure', async () => {
    mockAuth.mockReturnValueOnce({ ok: false, error: { code: 'AUTH_REQUIRED' } } as never);
    const res = await listTaskLists();
    expect(res.ok).toBe(false);
    expect(mockHttp).not.toHaveBeenCalled();
  });

  it('propagates http failure', async () => {
    mockHttp.mockResolvedValueOnce({ ok: false, error: { code: 'API_ERROR' } } as never);
    const res = await listTaskLists();
    expect(res.ok).toBe(false);
  });
});

describe('listTasks', () => {
  it('defaults to the default list and excludes completed', async () => {
    mockHttp
      .mockResolvedValueOnce(httpOk({ value: [{ id: 'def', displayName: 'Tasks', wellknownListName: 'defaultList' }] }))
      .mockResolvedValueOnce(httpOk({ value: [
        { id: 't1', title: 'A', status: 'notStarted', dueDateTime: { dateTime: '2026-06-10T00:00:00Z' }, body: { content: '<b>note</b>' } },
      ] }));
    const res = await listTasks({});
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.listId).toBe('def');
    expect(res.value.tasks[0]).toMatchObject({ id: 't1', title: 'A', dueDateTime: '2026-06-10T00:00:00Z', body: 'note' });
    const url = mockHttp.mock.calls[1][0] as string;
    expect(url).toContain('completed');
  });

  it('uses the given listId and can include completed', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ value: [] }));
    const res = await listTasks({ listId: 'L9', includeCompleted: true });
    expect(res.ok).toBe(true);
    const url = mockHttp.mock.calls[0][0] as string;
    expect(url).toContain('/me/todo/lists/L9/tasks');
    expect(url).not.toContain('status%20ne');
  });
});

describe('createTask', () => {
  it('posts a task with optional fields to the default list', async () => {
    mockHttp
      .mockResolvedValueOnce(httpOk({ value: [{ id: 'def', wellknownListName: 'defaultList' }] }))
      .mockResolvedValueOnce(httpOk({ id: 'new', title: 'Buy milk', status: 'notStarted' }));
    const res = await createTask({ title: 'Buy milk', dueDateTime: '2026-06-10T00:00:00Z', body: 'note', importance: 'high' });
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value).toMatchObject({ id: 'new', title: 'Buy milk' });
    const body = JSON.parse((mockHttp.mock.calls[1][1] as { body: string }).body);
    expect(body).toMatchObject({ title: 'Buy milk', importance: 'high' });
    expect(body.dueDateTime).toEqual({ dateTime: '2026-06-10T00:00:00Z', timeZone: 'UTC' });
    expect(body.body).toEqual({ content: 'note', contentType: 'text' });
  });
});

describe('createTask edge cases', () => {
  it('uses an explicit listId and omits unset optional fields', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ id: 'n', title: 'T', status: 'notStarted' }));
    const res = await createTask({ title: 'T', listId: 'L5' });
    expect(res.ok).toBe(true);
    const call = mockHttp.mock.calls[0];
    expect(call[0]).toContain('/me/todo/lists/L5/tasks');
    const body = JSON.parse((call[1] as { body: string }).body);
    expect(body).toEqual({ title: 'T' });
  });

  it('propagates auth failure without calling http', async () => {
    mockAuth.mockReturnValueOnce({ ok: false, error: { code: 'AUTH_REQUIRED' } } as never);
    const res = await createTask({ title: 'T', listId: 'L5' });
    expect(res.ok).toBe(false);
    expect(mockHttp).not.toHaveBeenCalled();
  });
});

describe('listTasks auth failure', () => {
  it('returns the auth error', async () => {
    mockAuth.mockReturnValueOnce({ ok: false, error: { code: 'AUTH_REQUIRED' } } as never);
    const res = await listTasks({ listId: 'L1' });
    expect(res.ok).toBe(false);
  });
});

describe('setTaskCompleted', () => {
  it('patches status to completed', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ id: 't1', title: 'A', status: 'completed' }));
    const res = await setTaskCompleted('L1', 't1', true);
    expect(res.ok).toBe(true);
    const call = mockHttp.mock.calls[0];
    expect(call[1]).toMatchObject({ method: 'PATCH' });
    expect(JSON.parse((call[1] as { body: string }).body)).toEqual({ status: 'completed' });
  });

  it('patches status to notStarted when reopening', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ id: 't1', status: 'notStarted' }));
    await setTaskCompleted('L1', 't1', false);
    expect(JSON.parse((mockHttp.mock.calls[0][1] as { body: string }).body)).toEqual({ status: 'notStarted' });
  });
});
