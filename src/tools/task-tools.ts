/**
 * Microsoft To Do task tool handlers.
 */

import { z } from 'zod';
import type { Tool } from '@modelcontextprotocol/sdk/types.js';
import type { RegisteredTool, ToolContext, ToolResult } from './index.js';
import { handleApiResult } from './index.js';
import {
  listTaskLists,
  listTasks,
  createTask,
  setTaskCompleted,
} from '../api/tasks-api.js';

// ─────────────────────────────────────────────────────────────────────────────
// Schemas
// ─────────────────────────────────────────────────────────────────────────────

export const ListTaskListsInputSchema = z.object({});

export const ListTasksInputSchema = z.object({
  listId: z.string().optional(),
  includeCompleted: z.boolean().optional().default(false),
  top: z.number().min(1).max(100).optional().default(50),
});

export const CreateTaskInputSchema = z.object({
  title: z.string().min(1),
  listId: z.string().optional(),
  dueDateTime: z.string().optional(),
  body: z.string().optional(),
  importance: z.enum(['low', 'normal', 'high']).optional(),
});

export const CompleteTaskInputSchema = z.object({
  listId: z.string().min(1),
  taskId: z.string().min(1),
  completed: z.boolean().optional().default(true),
});

// ─────────────────────────────────────────────────────────────────────────────
// Definitions
// ─────────────────────────────────────────────────────────────────────────────

const listTaskListsToolDefinition: Tool = {
  name: 'teams_list_task_lists',
  description: 'List the user\'s Microsoft To Do task lists (the default list plus any custom lists). Use the returned list id with teams_list_tasks or teams_create_task.',
  inputSchema: { type: 'object', properties: {} },
};

const listTasksToolDefinition: Tool = {
  name: 'teams_list_tasks',
  description: 'List the user\'s Microsoft To Do tasks. Defaults to the default list and to open (not completed) tasks. Pass a listId from teams_list_task_lists to target a specific list.',
  inputSchema: {
    type: 'object',
    properties: {
      listId: { type: 'string', description: 'Task list id (from teams_list_task_lists). Defaults to the default list.' },
      includeCompleted: { type: 'boolean', description: 'Include completed tasks (default false).' },
      top: { type: 'number', description: 'Max tasks to return (default 50).' },
    },
  },
};

const createTaskToolDefinition: Tool = {
  name: 'teams_create_task',
  description: 'Create a Microsoft To Do task. Goes into the default list unless a listId is given. Optionally set a due date (ISO 8601 UTC), a note body, and importance.',
  inputSchema: {
    type: 'object',
    properties: {
      title: { type: 'string', description: 'Task title.' },
      listId: { type: 'string', description: 'Task list id (defaults to the default list).' },
      dueDateTime: { type: 'string', description: 'Optional due date/time (ISO 8601 UTC).' },
      body: { type: 'string', description: 'Optional note/description.' },
      importance: { type: 'string', enum: ['low', 'normal', 'high'], description: 'Optional importance.' },
    },
    required: ['title'],
  },
};

const completeTaskToolDefinition: Tool = {
  name: 'teams_complete_task',
  description: 'Mark a Microsoft To Do task as completed (or reopen it with completed=false). Needs the listId and taskId (from teams_list_tasks).',
  inputSchema: {
    type: 'object',
    properties: {
      listId: { type: 'string', description: 'The task list id.' },
      taskId: { type: 'string', description: 'The task id to update.' },
      completed: { type: 'boolean', description: 'True to complete (default), false to reopen.' },
    },
    required: ['listId', 'taskId'],
  },
};

// ─────────────────────────────────────────────────────────────────────────────
// Handlers
// ─────────────────────────────────────────────────────────────────────────────

async function handleListTaskLists(_input: z.infer<typeof ListTaskListsInputSchema>, _ctx: ToolContext): Promise<ToolResult> {
  return handleApiResult(await listTaskLists(), (v) => ({ lists: v.lists }));
}

async function handleListTasks(input: z.infer<typeof ListTasksInputSchema>, _ctx: ToolContext): Promise<ToolResult> {
  return handleApiResult(
    await listTasks({ listId: input.listId, includeCompleted: input.includeCompleted, top: input.top }),
    (v) => ({ listId: v.listId, count: v.tasks.length, tasks: v.tasks })
  );
}

async function handleCreateTask(input: z.infer<typeof CreateTaskInputSchema>, _ctx: ToolContext): Promise<ToolResult> {
  return handleApiResult(
    await createTask({ title: input.title, listId: input.listId, dueDateTime: input.dueDateTime, body: input.body, importance: input.importance }),
    (v) => ({ task: v, message: 'Task created.' })
  );
}

async function handleCompleteTask(input: z.infer<typeof CompleteTaskInputSchema>, _ctx: ToolContext): Promise<ToolResult> {
  return handleApiResult(
    await setTaskCompleted(input.listId, input.taskId, input.completed),
    (v) => ({ task: v, message: input.completed ? 'Task completed.' : 'Task reopened.' })
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// Exports
// ─────────────────────────────────────────────────────────────────────────────

export const listTaskListsTool: RegisteredTool<typeof ListTaskListsInputSchema> = {
  definition: listTaskListsToolDefinition, schema: ListTaskListsInputSchema, handler: handleListTaskLists,
};
export const listTasksTool: RegisteredTool<typeof ListTasksInputSchema> = {
  definition: listTasksToolDefinition, schema: ListTasksInputSchema, handler: handleListTasks,
};
export const createTaskTool: RegisteredTool<typeof CreateTaskInputSchema> = {
  definition: createTaskToolDefinition, schema: CreateTaskInputSchema, handler: handleCreateTask,
};
export const completeTaskTool: RegisteredTool<typeof CompleteTaskInputSchema> = {
  definition: completeTaskToolDefinition, schema: CompleteTaskInputSchema, handler: handleCompleteTask,
};

export const taskTools = [listTaskListsTool, listTasksTool, createTaskTool, completeTaskTool];
