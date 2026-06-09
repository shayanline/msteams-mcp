/**
 * Channel management tool handlers.
 */

import { z } from 'zod';
import type { Tool } from '@modelcontextprotocol/sdk/types.js';
import type { RegisteredTool, ToolContext, ToolResult } from './index.js';
import { createChannel, deleteChannel } from '../api/channels-api.js';
import { handleApiResult } from './index.js';

// ─────────────────────────────────────────────────────────────────────────────
// Schemas
// ─────────────────────────────────────────────────────────────────────────────

export const CreateChannelInputSchema = z.object({
  teamId: z.string().min(1, 'Team ID cannot be empty'),
  displayName: z.string().min(1).max(50),
  description: z.string().optional(),
  membershipType: z.enum(['standard', 'private']).optional().default('standard'),
});

export const DeleteChannelInputSchema = z.object({
  teamId: z.string().min(1, 'Team ID cannot be empty'),
  channelId: z.string().min(1, 'Channel ID cannot be empty'),
});

// ─────────────────────────────────────────────────────────────────────────────
// Tool Definitions
// ─────────────────────────────────────────────────────────────────────────────

const createChannelDefinition: Tool = {
  name: 'teams_create_channel',
  description: 'Create a new channel in a Teams team. Provide the team\'s group ID (teamId, from teams_find_channel results) and a channel name. Returns the new channel\'s conversationId, which can be used with teams_send_message, teams_send_file, etc. This changes a real team, so confirm with the user before creating.',
  inputSchema: {
    type: 'object' as const,
    properties: {
      teamId: {
        type: 'string',
        description: 'The team group ID (GUID). Get this from teams_find_channel results (teamId field).',
      },
      displayName: {
        type: 'string',
        description: 'The channel name (max 50 characters).',
      },
      description: {
        type: 'string',
        description: 'Optional channel description.',
      },
      membershipType: {
        type: 'string',
        enum: ['standard', 'private'],
        description: 'Channel type: "standard" (default, visible to all team members) or "private".',
      },
    },
    required: ['teamId', 'displayName'],
  },
};

const deleteChannelDefinition: Tool = {
  name: 'teams_delete_channel',
  description: 'Delete a channel from a Teams team. Provide the team\'s group ID (teamId, from teams_find_channel) and the channel\'s conversationId. This permanently removes the channel and its content for everyone, so confirm with the user before deleting. The team\'s General channel cannot be deleted.',
  inputSchema: {
    type: 'object' as const,
    properties: {
      teamId: {
        type: 'string',
        description: 'The team group ID (GUID). Get this from teams_find_channel results (teamId field).',
      },
      channelId: {
        type: 'string',
        description: 'The channel conversation ID to delete (19:...@thread.tacv2).',
      },
    },
    required: ['teamId', 'channelId'],
  },
};

// ─────────────────────────────────────────────────────────────────────────────
// Handlers
// ─────────────────────────────────────────────────────────────────────────────

async function handleCreateChannel(
  input: z.infer<typeof CreateChannelInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await createChannel(input.teamId, input.displayName, input.description, input.membershipType);
  return handleApiResult(result, (value) => ({
    success: true,
    ...value,
    message: `Channel "${value.displayName}" created.`,
  }));
}

async function handleDeleteChannel(
  input: z.infer<typeof DeleteChannelInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await deleteChannel(input.teamId, input.channelId);
  return handleApiResult(result, (value) => ({
    success: true,
    ...value,
    message: 'Channel deleted.',
  }));
}

// ─────────────────────────────────────────────────────────────────────────────
// Registration
// ─────────────────────────────────────────────────────────────────────────────

const createChannelTool: RegisteredTool<typeof CreateChannelInputSchema> = {
  definition: createChannelDefinition,
  schema: CreateChannelInputSchema,
  handler: handleCreateChannel,
};

const deleteChannelTool: RegisteredTool<typeof DeleteChannelInputSchema> = {
  definition: deleteChannelDefinition,
  schema: DeleteChannelInputSchema,
  handler: handleDeleteChannel,
};

export const channelTools = [createChannelTool, deleteChannelTool];
