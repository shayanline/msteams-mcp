/**
 * Microsoft To Do tasks API client (Microsoft Graph).
 *
 * Uses the Graph token (Tasks.ReadWrite scope) extracted from the Teams session.
 */

import { httpRequest } from '../utils/http.js';
import { GRAPH_BASE_URL, getGraphHeaders } from '../utils/api-config.js';
import { type Result, ok } from '../types/result.js';
import { requireGraphAuth } from '../utils/auth-guards.js';

// ─────────────────────────────────────────────────────────────────────────────
// Types
// ─────────────────────────────────────────────────────────────────────────────

export interface TaskList {
  id: string;
  displayName: string;
  isDefault: boolean;
}

export interface TodoTask {
  id: string;
  title: string;
  status: string;
  importance?: string;
  dueDateTime?: string;
  body?: string;
  createdDateTime?: string;
}

// ─────────────────────────────────────────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────────────────────────────────────────

function parseTask(raw: Record<string, unknown>): TodoTask {
  const due = raw.dueDateTime as { dateTime?: string } | undefined;
  const body = raw.body as { content?: string } | undefined;
  return {
    id: raw.id as string,
    title: (raw.title as string) ?? '',
    status: (raw.status as string) ?? 'notStarted',
    importance: raw.importance as string | undefined,
    dueDateTime: due?.dateTime,
    body: body?.content ? body.content.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim() || undefined : undefined,
    createdDateTime: raw.createdDateTime as string | undefined,
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// API
// ─────────────────────────────────────────────────────────────────────────────

/** Lists the user's To Do task lists. */
export async function listTaskLists(): Promise<Result<{ lists: TaskList[] }>> {
  const auth = requireGraphAuth();
  if (!auth.ok) return auth;

  const response = await httpRequest<{ value: Array<Record<string, unknown>> }>(
    `${GRAPH_BASE_URL}/me/todo/lists`,
    { method: 'GET', headers: getGraphHeaders(auth.value) }
  );
  if (!response.ok) return response;

  const lists = (response.value.data.value ?? []).map((l) => ({
    id: l.id as string,
    displayName: (l.displayName as string) ?? '',
    isDefault: l.wellknownListName === 'defaultList',
  }));
  return ok({ lists });
}

/** Resolves the default task list id (or the first list). */
async function defaultListId(): Promise<Result<string>> {
  const lists = await listTaskLists();
  if (!lists.ok) return lists;
  const def = lists.value.lists.find((l) => l.isDefault) ?? lists.value.lists[0];
  return ok(def?.id ?? '');
}

/** Lists tasks in a list (defaults to the default list). Excludes completed unless includeCompleted. */
export async function listTasks(
  options: { listId?: string; includeCompleted?: boolean; top?: number } = {}
): Promise<Result<{ listId: string; tasks: TodoTask[] }>> {
  const auth = requireGraphAuth();
  if (!auth.ok) return auth;

  let listId = options.listId;
  if (!listId) {
    const def = await defaultListId();
    if (!def.ok) return def;
    listId = def.value;
  }

  const params = new URLSearchParams({ $top: String(options.top ?? 50), $orderby: 'createdDateTime desc' });
  if (!options.includeCompleted) params.set('$filter', "status ne 'completed'");

  const response = await httpRequest<{ value: Array<Record<string, unknown>> }>(
    `${GRAPH_BASE_URL}/me/todo/lists/${encodeURIComponent(listId)}/tasks?${params.toString()}`,
    { method: 'GET', headers: getGraphHeaders(auth.value) }
  );
  if (!response.ok) return response;

  return ok({ listId, tasks: (response.value.data.value ?? []).map(parseTask) });
}

/** Creates a task in a list (defaults to the default list). */
export async function createTask(
  options: { title: string; listId?: string; dueDateTime?: string; body?: string; importance?: 'low' | 'normal' | 'high' }
): Promise<Result<TodoTask>> {
  const auth = requireGraphAuth();
  if (!auth.ok) return auth;

  let listId = options.listId;
  if (!listId) {
    const def = await defaultListId();
    if (!def.ok) return def;
    listId = def.value;
  }

  const taskBody: Record<string, unknown> = { title: options.title };
  if (options.importance) taskBody.importance = options.importance;
  if (options.body) taskBody.body = { content: options.body, contentType: 'text' };
  if (options.dueDateTime) taskBody.dueDateTime = { dateTime: options.dueDateTime, timeZone: 'UTC' };

  const response = await httpRequest<Record<string, unknown>>(
    `${GRAPH_BASE_URL}/me/todo/lists/${encodeURIComponent(listId)}/tasks`,
    { method: 'POST', headers: getGraphHeaders(auth.value), body: JSON.stringify(taskBody) }
  );
  if (!response.ok) return response;
  return ok(parseTask(response.value.data));
}

/** Marks a task complete (or reopens it). */
export async function setTaskCompleted(
  listId: string,
  taskId: string,
  completed = true
): Promise<Result<TodoTask>> {
  const auth = requireGraphAuth();
  if (!auth.ok) return auth;

  const response = await httpRequest<Record<string, unknown>>(
    `${GRAPH_BASE_URL}/me/todo/lists/${encodeURIComponent(listId)}/tasks/${encodeURIComponent(taskId)}`,
    {
      method: 'PATCH',
      headers: getGraphHeaders(auth.value),
      body: JSON.stringify({ status: completed ? 'completed' : 'notStarted' }),
    }
  );
  if (!response.ok) return response;
  return ok(parseTask(response.value.data));
}
