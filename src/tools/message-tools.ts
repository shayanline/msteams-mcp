/**
 * Messaging-related tool handlers.
 */

import { z } from 'zod';
import type { Tool } from '@modelcontextprotocol/sdk/types.js';
import type { RegisteredTool, ToolContext, ToolResult } from './index.js';
import {
  sendMessage,
  getMessage,
  saveMessage,
  unsaveMessage,
  getOneOnOneChatId,
  createGroupChat,
  editMessage,
  deleteMessage,
  getUnreadStatus,
  getUnreadConversations,
  markAsRead,
  getActivityFeed,
  addReaction,
  removeReaction,
  getSavedMessages,
  getFollowedThreads,
  listConversations,
  markUnread,
  getChatMembers,
  addMember,
  removeMember,
  leaveChat,
  renameChat,
  forwardMessage,
  pinMessage,
  unpinMessage,
  setMuted,
} from '../api/chatsvc-api.js';
import { getFavorites, addFavorite, removeFavorite, getCustomEmojis, getMyTeamsAndChannels } from '../api/csa-api.js';
import { SELF_CHAT_ID, MAX_THREAD_LIMIT, STANDARD_EMOJIS } from '../constants.js';

// ─────────────────────────────────────────────────────────────────────────────
// Schemas
// ─────────────────────────────────────────────────────────────────────────────

export const SendMessageInputSchema = z.object({
  content: z.string().min(1, 'Message content cannot be empty'),
  conversationId: z.string().optional().default(SELF_CHAT_ID),
  replyToMessageId: z.string().optional(),
  importance: z.enum(['normal', 'high', 'urgent']).optional(),
  subject: z.string().optional(),
});

export const FavoriteInputSchema = z.object({
  conversationId: z.string().min(1, 'Conversation ID cannot be empty'),
});

export const SaveMessageInputSchema = z.object({
  conversationId: z.string().min(1, 'Conversation ID cannot be empty'),
  messageId: z.string().min(1, 'Message ID cannot be empty'),
  rootMessageId: z.string().optional(),
});

export const GetChatInputSchema = z.object({
  userId: z.string().min(1, 'User ID cannot be empty'),
});

export const CreateGroupChatInputSchema = z.object({
  userIds: z.array(z.string().min(1)).min(2, 'At least 2 user IDs are required for a group chat'),
  topic: z.string().optional(),
});

export const EditMessageInputSchema = z.object({
  conversationId: z.string().min(1, 'Conversation ID cannot be empty'),
  messageId: z.string().min(1, 'Message ID cannot be empty'),
  content: z.string().min(1, 'Content cannot be empty'),
});

export const DeleteMessageInputSchema = z.object({
  conversationId: z.string().min(1, 'Conversation ID cannot be empty'),
  messageId: z.string().min(1, 'Message ID cannot be empty'),
});

export const GetUnreadInputSchema = z.object({
  conversationId: z.string().min(1).optional(),
});

export const MarkAsReadInputSchema = z.object({
  conversationId: z.string().min(1, 'Conversation ID cannot be empty'),
  messageId: z.string().min(1, 'Message ID cannot be empty'),
});

export const GetActivityInputSchema = z.object({
  limit: z.number().min(1).max(200).optional(),
  syncState: z.string().optional(),
});

export const SearchEmojiInputSchema = z.object({
  query: z.string().min(1, 'Search query cannot be empty'),
});

export const AddReactionInputSchema = z.object({
  conversationId: z.string().min(1, 'Conversation ID cannot be empty'),
  messageId: z.string().min(1, 'Message ID cannot be empty'),
  emoji: z.string().min(1, 'Emoji key cannot be empty'),
});

export const RemoveReactionInputSchema = z.object({
  conversationId: z.string().min(1, 'Conversation ID cannot be empty'),
  messageId: z.string().min(1, 'Message ID cannot be empty'),
  emoji: z.string().min(1, 'Emoji key cannot be empty'),
});

export const GetSavedMessagesInputSchema = z.object({
  limit: z.number().min(1).max(MAX_THREAD_LIMIT).optional(),
});

export const GetFollowedThreadsInputSchema = z.object({
  limit: z.number().min(1).max(MAX_THREAD_LIMIT).optional(),
});

export const GetMessageInputSchema = z.object({
  conversationId: z.string().min(1, 'Conversation ID cannot be empty'),
  messageId: z.string().min(1, 'Message ID cannot be empty'),
});

export const ListChatsInputSchema = z.object({
  limit: z.number().min(1).max(200).optional().default(50),
  type: z.enum(['chat', 'channel', 'meeting']).optional(),
});

export const MarkUnreadInputSchema = z.object({
  conversationId: z.string().min(1, 'Conversation ID cannot be empty'),
  messageId: z.string().min(1, 'Message ID cannot be empty'),
});

export const ListTeamsInputSchema = z.object({});

export const GetChatMembersInputSchema = z.object({
  conversationId: z.string().min(1, 'Conversation ID cannot be empty'),
});

export const AddMemberInputSchema = z.object({
  conversationId: z.string().min(1, 'Conversation ID cannot be empty'),
  userId: z.string().min(1, 'User ID cannot be empty'),
  role: z.enum(['Admin', 'User']).optional().default('User'),
});

export const RemoveMemberInputSchema = z.object({
  conversationId: z.string().min(1, 'Conversation ID cannot be empty'),
  userId: z.string().min(1, 'User ID cannot be empty'),
});

export const LeaveChatInputSchema = z.object({
  conversationId: z.string().min(1, 'Conversation ID cannot be empty'),
});

export const RenameChatInputSchema = z.object({
  conversationId: z.string().min(1, 'Conversation ID cannot be empty'),
  topic: z.string().min(1, 'Topic cannot be empty'),
});

export const ForwardMessageInputSchema = z.object({
  sourceConversationId: z.string().min(1, 'Source conversation ID cannot be empty'),
  messageId: z.string().min(1, 'Message ID cannot be empty'),
  targetConversationId: z.string().min(1, 'Target conversation ID cannot be empty'),
  comment: z.string().optional(),
});

export const PinMessageInputSchema = z.object({
  conversationId: z.string().min(1, 'Conversation ID cannot be empty'),
  messageId: z.string().min(1, 'Message ID cannot be empty'),
});

export const UnpinMessageInputSchema = z.object({
  conversationId: z.string().min(1, 'Conversation ID cannot be empty'),
});

export const MuteChatInputSchema = z.object({
  conversationId: z.string().min(1, 'Conversation ID cannot be empty'),
});

// ─────────────────────────────────────────────────────────────────────────────
// Tool Definitions
// ─────────────────────────────────────────────────────────────────────────────

const sendMessageToolDefinition: Tool = {
  name: 'teams_send_message',
  description: 'Send a message to a Teams conversation. Use plain markdown (not HTML): **bold**, *italic*, ~~strikethrough~~, `code`, ```code blocks```, lists, headings, tables, blockquotes. For spacing, just write normal markdown: a blank line between paragraphs renders as a visible gap and a single newline is a line break (no special escaping or trailing-space tricks needed). Supports @mentions: people with @[Name](mri) (MRI from teams_search_people) and channel tags with @[TagName](tag:tagId) (IDs from teams_get_tags). Defaults to self-notes (48:notes). For channel thread replies, provide replyToMessageId.',
  inputSchema: {
    type: 'object',
    properties: {
      content: {
        type: 'string',
        description: 'The message content in markdown (not HTML). Supports: **bold**, *italic*, ~~strikethrough~~, `inline code`, ```code blocks```, bullet lists (- item), numbered lists (1. item), headings, pipe tables and blockquotes. Spacing works like normal markdown: a blank line between blocks becomes a visible paragraph gap, a single newline is a line break. Just write it naturally; do not add trailing-space hard breaks or raw HTML tags. For people @mentions use @[DisplayName](mri) (MRI from teams_search_people). For channel tag @mentions use @[DisplayName](tag:tagId) (tag IDs from teams_get_tags). Markdown links [text](url) are supported.',
      },
      conversationId: {
        type: 'string',
        description: 'The conversation ID to send to. Use "48:notes" for self-chat (default), or a channel/chat conversation ID.',
      },
      replyToMessageId: {
        type: 'string',
        description: 'For channel thread replies: the message ID of the thread root. Use serverMessageId from teams_send_message, id from teams_get_thread, or messageId from teams_search.',
      },
      importance: {
        type: 'string',
        enum: ['normal', 'high', 'urgent'],
        description: 'Message priority. "high" marks it important, "urgent" sends a priority notification that repeats until read. Defaults to normal. Use urgent sparingly.',
      },
      subject: {
        type: 'string',
        description: 'Optional subject/title. Shown as the bold heading on a channel post. Ignored for ordinary chats.',
      },
    },
    required: ['content'],
  },
};

const getFavoritesToolDefinition: Tool = {
  name: 'teams_get_favorites',
  description: 'Get the user\'s favourite/pinned conversations in Teams. Returns conversation IDs with display names (channel name, chat topic, or participant names) and type (Channel, Chat, Meeting).',
  inputSchema: {
    type: 'object',
    properties: {},
  },
};

const addFavoriteToolDefinition: Tool = {
  name: 'teams_add_favorite',
  description: 'Add a conversation to the user\'s favourites/pinned list.',
  inputSchema: {
    type: 'object',
    properties: {
      conversationId: {
        type: 'string',
        description: 'The conversation ID to pin (e.g., "19:abc@thread.tacv2")',
      },
    },
    required: ['conversationId'],
  },
};

const removeFavoriteToolDefinition: Tool = {
  name: 'teams_remove_favorite',
  description: 'Remove a conversation from the user\'s favourites/pinned list.',
  inputSchema: {
    type: 'object',
    properties: {
      conversationId: {
        type: 'string',
        description: 'The conversation ID to unpin',
      },
    },
    required: ['conversationId'],
  },
};

const saveMessageToolDefinition: Tool = {
  name: 'teams_save_message',
  description: 'Save (bookmark) a message in Teams. Saved messages can be accessed later from the Saved view in Teams.',
  inputSchema: {
    type: 'object',
    properties: {
      conversationId: {
        type: 'string',
        description: 'The conversation ID containing the message',
      },
      messageId: {
        type: 'string',
        description: 'The message ID to save. Use: serverMessageId from teams_send_message, id from teams_get_thread, or messageId from teams_search.',
      },
      rootMessageId: {
        type: 'string',
        description: 'For channel threaded replies only: the ID of the thread root post. Not needed for top-level posts or non-channel conversations.',
      },
    },
    required: ['conversationId', 'messageId'],
  },
};

const unsaveMessageToolDefinition: Tool = {
  name: 'teams_unsave_message',
  description: 'Remove a saved (bookmarked) message in Teams.',
  inputSchema: {
    type: 'object',
    properties: {
      conversationId: {
        type: 'string',
        description: 'The conversation ID containing the message',
      },
      messageId: {
        type: 'string',
        description: 'The message ID to unsave',
      },
      rootMessageId: {
        type: 'string',
        description: 'For channel threaded replies only: the ID of the thread root post. Not needed for top-level posts or non-channel conversations.',
      },
    },
    required: ['conversationId', 'messageId'],
  },
};

const getChatToolDefinition: Tool = {
  name: 'teams_get_chat',
  description: 'Get the conversation ID for a 1:1 chat with a person. Use this to start a new chat or find an existing one. The conversation ID can then be used with teams_send_message to send messages.',
  inputSchema: {
    type: 'object',
    properties: {
      userId: {
        type: 'string',
        description: 'The user\'s identifier. Can be: MRI (8:orgid:guid), object ID with tenant (guid@tenantId), or raw object ID (guid). Get this from teams_search_people results.',
      },
    },
    required: ['userId'],
  },
};

const createGroupChatToolDefinition: Tool = {
  name: 'teams_create_group_chat',
  description: 'Create a new group chat with multiple people. Returns a conversation ID for use with teams_send_message. You are automatically included as a member. For 1:1 chats, use teams_get_chat instead.',
  inputSchema: {
    type: 'object',
    properties: {
      userIds: {
        type: 'array',
        items: { type: 'string' },
        description: 'Array of OTHER users to add (at least 2, not including yourself). Can be: MRI (8:orgid:guid), object ID with tenant (guid@tenantId), or raw GUID. Get these from teams_search_people or teams_get_frequent_contacts.',
        minItems: 2,
      },
      topic: {
        type: 'string',
        description: 'Optional chat name/topic. If omitted, Teams shows member names.',
      },
    },
    required: ['userIds'],
  },
};

const editMessageToolDefinition: Tool = {
  name: 'teams_edit_message',
  description: 'Edit one of your own messages (same markdown and @mention rules as teams_send_message: people @[Name](mri), channel tags @[TagName](tag:tagId) from teams_get_tags). You can only edit messages you sent. WARNING: editing replaces the message body and DROPS any file attachments on it. To change a message that has a file, delete it (teams_delete_message) and re-send with teams_send_file instead of editing.',
  inputSchema: {
    type: 'object',
    properties: {
      conversationId: {
        type: 'string',
        description: 'The conversation ID containing the message',
      },
      messageId: {
        type: 'string',
        description: 'The message ID to edit. Use: serverMessageId from teams_send_message, id from teams_get_thread, or messageId from teams_search.',
      },
      content: {
        type: 'string',
        description: 'New content in markdown (not raw HTML): **bold**, *italic*, lists, code, @[Person](mri), @[Tag](tag:id), [text](url) — same as teams_send_message.',
      },
    },
    required: ['conversationId', 'messageId', 'content'],
  },
};

const deleteMessageToolDefinition: Tool = {
  name: 'teams_delete_message',
  description: 'Delete one of your own messages (soft delete - the message remains but content becomes empty). You can only delete messages you sent, unless you are a channel owner/moderator.',
  inputSchema: {
    type: 'object',
    properties: {
      conversationId: {
        type: 'string',
        description: 'The conversation ID containing the message',
      },
      messageId: {
        type: 'string',
        description: 'The message ID to delete. Use: serverMessageId from teams_send_message, id from teams_get_thread, or messageId from teams_search.',
      },
    },
    required: ['conversationId', 'messageId'],
  },
};

const getUnreadToolDefinition: Tool = {
  name: 'teams_get_unread',
  description: 'Get unread status. Without conversationId: one bulk API call over your recent conversations (up to 200), returns separate lists of unread chats and channels (conversationId, displayName, lastMessageFrom) plus counts. With conversationId: unread count for that chat/channel using read horizon vs recent messages.',
  inputSchema: {
    type: 'object',
    properties: {
      conversationId: {
        type: 'string',
        description: 'Optional. If set, returns unreadCount (and related fields) for this conversation only. If omitted, returns bulk unreadChats/unreadChannels across recent conversations.',
      },
    },
  },
};

const markAsReadToolDefinition: Tool = {
  name: 'teams_mark_read',
  description: 'Mark a conversation as read up to a specific message. This updates your read position so messages up to (and including) the specified message are marked as read.',
  inputSchema: {
    type: 'object',
    properties: {
      conversationId: {
        type: 'string',
        description: 'The conversation ID to mark as read',
      },
      messageId: {
        type: 'string',
        description: 'The message ID to mark as read up to (all messages up to this point will be marked read)',
      },
    },
    required: ['conversationId', 'messageId'],
  },
};

const getActivityToolDefinition: Tool = {
  name: 'teams_get_activity',
  description: 'Get the user\'s activity feed - mentions, reactions, replies, and other notifications. Returns recent activity items with sender, content, and source conversation context. Pass syncState from a previous response to get only newer items.',
  inputSchema: {
    type: 'object',
    properties: {
      limit: {
        type: 'number',
        description: 'Maximum number of activity items to return (default: 50, max: 200)',
      },
      syncState: {
        type: 'string',
        description: 'Pagination token from a previous teams_get_activity response. When provided, returns only activity newer than the previous fetch.',
      },
    },
  },
};

const searchEmojiToolDefinition: Tool = {
  name: 'teams_search_emoji',
  description: 'Search for emojis by name or keyword. Returns both standard Teams emojis and custom organisation emojis, indicating which is which. Use the returned key with teams_add_reaction.',
  inputSchema: {
    type: 'object',
    properties: {
      query: {
        type: 'string',
        description: 'Search term (e.g., "thumbs", "heart", "laugh", "cat")',
      },
    },
    required: ['query'],
  },
};

const addReactionToolDefinition: Tool = {
  name: 'teams_add_reaction',
  description: 'Add an emoji reaction to a message. Common reactions: like (👍), heart (❤️), laugh (😂), surprised (😮), sad (😢), angry (😠). Use teams_search_emoji to find other emojis.',
  inputSchema: {
    type: 'object',
    properties: {
      conversationId: {
        type: 'string',
        description: 'The conversation ID containing the message',
      },
      messageId: {
        type: 'string',
        description: 'The message ID to react to. Use: serverMessageId from teams_send_message, id from teams_get_thread, or messageId from teams_search. NOT the messageId from teams_send_message (that is client-generated and will fail).',
      },
      emoji: {
        type: 'string',
        description: 'The emoji key (e.g., "like", "heart", "laugh"). Get from teams_search_emoji or use common ones directly.',
      },
    },
    required: ['conversationId', 'messageId', 'emoji'],
  },
};

const removeReactionToolDefinition: Tool = {
  name: 'teams_remove_reaction',
  description: 'Remove an emoji reaction from a message.',
  inputSchema: {
    type: 'object',
    properties: {
      conversationId: {
        type: 'string',
        description: 'The conversation ID containing the message',
      },
      messageId: {
        type: 'string',
        description: 'The message ID to remove the reaction from. Use: serverMessageId from teams_send_message, id from teams_get_thread, or messageId from teams_search.',
      },
      emoji: {
        type: 'string',
        description: 'The emoji key to remove (e.g., "like", "heart")',
      },
    },
    required: ['conversationId', 'messageId', 'emoji'],
  },
};

const getSavedMessagesToolDefinition: Tool = {
  name: 'teams_get_saved_messages',
  description: 'Get the list of messages the user has saved (bookmarked) in Teams. Returns references to saved messages with source conversation IDs and direct links. Use teams_get_message with sourceConversationId and sourceMessageId to fetch the full message content.',
  inputSchema: {
    type: 'object',
    properties: {
      limit: {
        type: 'number',
        description: 'Maximum number of saved messages to return (default: 50, max: 200)',
      },
    },
  },
};

const getFollowedThreadsToolDefinition: Tool = {
  name: 'teams_get_followed_threads',
  description: 'Get the list of threads the user is following in Teams. Returns references to followed threads with source conversation IDs and direct links. Use teams_get_message with sourceConversationId and sourcePostId to fetch the full post, or teams_get_thread for the full thread.',
  inputSchema: {
    type: 'object',
    properties: {
      limit: {
        type: 'number',
        description: 'Maximum number of followed threads to return (default: 50, max: 200)',
      },
    },
  },
};

const getMessageToolDefinition: Tool = {
  name: 'teams_get_message',
  description: 'Get a single message by ID with full content. Works for messages of any age - no retention limit. Use this to resolve truncated search results, saved message stubs, or retrieve any specific message when you have the conversationId and messageId.',
  inputSchema: {
    type: 'object',
    properties: {
      conversationId: {
        type: 'string',
        description: 'The conversation ID containing the message (e.g., from search results, saved messages, or activity feed)',
      },
      messageId: {
        type: 'string',
        description: 'The message ID to fetch (e.g., from search results, saved messages, or teams_get_thread)',
      },
    },
    required: ['conversationId', 'messageId'],
  },
};

const listChatsToolDefinition: Tool = {
  name: 'teams_list_chats',
  description: 'List the user\'s recent conversations (chats, group chats, channels, meetings) with display name/topic, type, favourite flag, last message preview and unread state. Use this to browse or find a conversation when you do not already have its ID. Optionally filter by type.',
  inputSchema: {
    type: 'object',
    properties: {
      limit: {
        type: 'number',
        description: 'Maximum conversations to return, most recent first (default 50, max 200).',
      },
      type: {
        type: 'string',
        enum: ['chat', 'channel', 'meeting'],
        description: 'Optional filter to only chats, only channels, or only meetings.',
      },
    },
  },
};

const markUnreadToolDefinition: Tool = {
  name: 'teams_mark_unread',
  description: 'Mark a conversation as unread from a specific message onward. Pass the message you want to become the first unread one. This moves the read marker back to just before that message.',
  inputSchema: {
    type: 'object',
    properties: {
      conversationId: {
        type: 'string',
        description: 'The conversation ID to mark unread.',
      },
      messageId: {
        type: 'string',
        description: 'The message ID that should become the first unread message (everything from here onward shows unread).',
      },
    },
    required: ['conversationId', 'messageId'],
  },
};

const listTeamsToolDefinition: Tool = {
  name: 'teams_list_teams',
  description: 'List all teams the user is a member of, each with its channels (id and name). Use this to browse teams and channels, or to get a channel\'s conversation ID for sending messages.',
  inputSchema: {
    type: 'object',
    properties: {},
  },
};

const getChatMembersToolDefinition: Tool = {
  name: 'teams_get_chat_members',
  description: 'List the members of a group chat or channel thread, with each member\'s MRI and role (Admin or User). Use teams_search_people to resolve an MRI to a name if needed.',
  inputSchema: {
    type: 'object',
    properties: {
      conversationId: { type: 'string', description: 'The group chat or channel conversation ID.' },
    },
    required: ['conversationId'],
  },
};

const addMemberToolDefinition: Tool = {
  name: 'teams_add_member',
  description: 'Add a person to a group chat. Confirm with the user before adding people. Works on group chats (not 1:1). Get the user identifier from teams_search_people.',
  inputSchema: {
    type: 'object',
    properties: {
      conversationId: { type: 'string', description: 'The group chat conversation ID.' },
      userId: { type: 'string', description: 'The person to add: MRI (8:orgid:guid), guid@tenant, or raw GUID (from teams_search_people).' },
      role: { type: 'string', enum: ['Admin', 'User'], description: 'Role to grant (default User).' },
    },
    required: ['conversationId', 'userId'],
  },
};

const removeMemberToolDefinition: Tool = {
  name: 'teams_remove_member',
  description: 'Remove a person from a group chat. This is destructive: confirm with the user first. To remove yourself, use teams_leave_chat instead.',
  inputSchema: {
    type: 'object',
    properties: {
      conversationId: { type: 'string', description: 'The group chat conversation ID.' },
      userId: { type: 'string', description: 'The person to remove: MRI, guid@tenant, or raw GUID.' },
    },
    required: ['conversationId', 'userId'],
  },
};

const leaveChatToolDefinition: Tool = {
  name: 'teams_leave_chat',
  description: 'Leave a group chat (removes you from it). Confirm with the user before leaving.',
  inputSchema: {
    type: 'object',
    properties: {
      conversationId: { type: 'string', description: 'The group chat conversation ID to leave.' },
    },
    required: ['conversationId'],
  },
};

const renameChatToolDefinition: Tool = {
  name: 'teams_rename_chat',
  description: 'Rename a group chat (set its topic/title). Applies to group chats, not 1:1 chats.',
  inputSchema: {
    type: 'object',
    properties: {
      conversationId: { type: 'string', description: 'The group chat conversation ID.' },
      topic: { type: 'string', description: 'The new chat name/topic.' },
    },
    required: ['conversationId', 'topic'],
  },
};

const forwardMessageToolDefinition: Tool = {
  name: 'teams_forward_message',
  description: 'Forward a message to another conversation. Re-posts the original as a quoted block (sender name + content) and carries file attachments as native chiclets with no re-upload. No label is injected by default — pass an optional comment to prepend a note above the quote. Get the source IDs from teams_search or teams_get_thread.',
  inputSchema: {
    type: 'object',
    properties: {
      sourceConversationId: { type: 'string', description: 'Conversation ID the original message is in.' },
      messageId: { type: 'string', description: 'The message ID to forward (serverMessageId / id / messageId).' },
      targetConversationId: { type: 'string', description: 'Conversation ID to forward the message to.' },
      comment: { type: 'string', description: 'Optional note to add above the forwarded message.' },
    },
    required: ['sourceConversationId', 'messageId', 'targetConversationId'],
  },
};

const pinMessageToolDefinition: Tool = {
  name: 'teams_pin_message',
  description: 'Pin a message in a conversation so it shows in the chat\'s pinned items. Get the messageId from teams_get_thread or teams_search (use the serverMessageId / id).',
  inputSchema: {
    type: 'object',
    properties: {
      conversationId: { type: 'string', description: 'The conversation containing the message.' },
      messageId: { type: 'string', description: 'The message id to pin.' },
    },
    required: ['conversationId', 'messageId'],
  },
};

const unpinMessageToolDefinition: Tool = {
  name: 'teams_unpin_message',
  description: 'Remove the pinned message(s) from a conversation.',
  inputSchema: {
    type: 'object',
    properties: {
      conversationId: { type: 'string', description: 'The conversation to clear pinned messages from.' },
    },
    required: ['conversationId'],
  },
};

const muteChatToolDefinition: Tool = {
  name: 'teams_mute_chat',
  description: 'Mute a conversation (turns off notifications/alerts for you). Use teams_unmute_chat to re-enable.',
  inputSchema: {
    type: 'object',
    properties: {
      conversationId: { type: 'string', description: 'The conversation to mute.' },
    },
    required: ['conversationId'],
  },
};

const unmuteChatToolDefinition: Tool = {
  name: 'teams_unmute_chat',
  description: 'Unmute a conversation (turns notifications/alerts back on for you).',
  inputSchema: {
    type: 'object',
    properties: {
      conversationId: { type: 'string', description: 'The conversation to unmute.' },
    },
    required: ['conversationId'],
  },
};

// ─────────────────────────────────────────────────────────────────────────────
// Handlers
// ─────────────────────────────────────────────────────────────────────────────

async function handleSendMessage(
  input: z.infer<typeof SendMessageInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await sendMessage(input.conversationId, input.content, {
    replyToMessageId: input.replyToMessageId,
    importance: input.importance,
    subject: input.subject,
  });

  if (!result.ok) {
    return { success: false, error: result.error };
  }

  // The timestamp is the server-assigned ID - use this for reactions, threading, edits, etc.
  const serverMessageId = result.value.timestamp ? String(result.value.timestamp) : undefined;

  const response: Record<string, unknown> = {
    messageId: result.value.messageId,
    timestamp: result.value.timestamp,
    conversationId: input.conversationId,
  };

  // Always include serverMessageId - this is the ID to use for reactions, edits, etc.
  if (serverMessageId) {
    response.serverMessageId = serverMessageId;
  }

  // Include replyToMessageId in response if this was a thread reply
  if (input.replyToMessageId) {
    response.replyToMessageId = input.replyToMessageId;
    response.note = 'Message posted as a reply to the thread. Use serverMessageId (not messageId) for reactions, edits, or threading.';
  } else if (serverMessageId) {
    response.note = 'Use serverMessageId (not messageId) for reactions, edits, or threading.';
  }

  return { success: true, data: response };
}

async function handleGetFavorites(
  _input: Record<string, never>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await getFavorites();

  if (!result.ok) {
    return { success: false, error: result.error };
  }

  return {
    success: true,
    data: {
      count: result.value.favorites.length,
      favorites: result.value.favorites,
    },
  };
}

async function handleAddFavorite(
  input: z.infer<typeof FavoriteInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await addFavorite(input.conversationId);

  if (!result.ok) {
    return { success: false, error: result.error };
  }

  return {
    success: true,
    data: {
      message: `Added ${input.conversationId} to favourites`,
    },
  };
}

async function handleRemoveFavorite(
  input: z.infer<typeof FavoriteInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await removeFavorite(input.conversationId);

  if (!result.ok) {
    return { success: false, error: result.error };
  }

  return {
    success: true,
    data: {
      message: `Removed ${input.conversationId} from favourites`,
    },
  };
}

async function handleSaveMessage(
  input: z.infer<typeof SaveMessageInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await saveMessage(input.conversationId, input.messageId, input.rootMessageId);

  if (!result.ok) {
    return { success: false, error: result.error };
  }

  return {
    success: true,
    data: {
      message: 'Message saved',
      conversationId: input.conversationId,
      messageId: input.messageId,
    },
  };
}

async function handleUnsaveMessage(
  input: z.infer<typeof SaveMessageInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await unsaveMessage(input.conversationId, input.messageId, input.rootMessageId);

  if (!result.ok) {
    return { success: false, error: result.error };
  }

  return {
    success: true,
    data: {
      message: 'Message unsaved',
      conversationId: input.conversationId,
      messageId: input.messageId,
    },
  };
}

async function handleGetChat(
  input: z.infer<typeof GetChatInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = getOneOnOneChatId(input.userId);

  if (!result.ok) {
    return { success: false, error: result.error };
  }

  return {
    success: true,
    data: {
      conversationId: result.value.conversationId,
      otherUserId: result.value.otherUserId,
      currentUserId: result.value.currentUserId,
      note: 'Use this conversationId with teams_send_message to send a message. The conversation is created automatically when the first message is sent.',
    },
  };
}

async function handleCreateGroupChat(
  input: z.infer<typeof CreateGroupChatInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await createGroupChat(input.userIds, input.topic);

  if (!result.ok) {
    return { success: false, error: result.error };
  }

  return {
    success: true,
    data: {
      conversationId: result.value.conversationId,
      members: result.value.members,
      topic: result.value.topic,
      note: 'Use this conversationId with teams_send_message to send messages to the group.',
    },
  };
}

async function handleEditMessage(
  input: z.infer<typeof EditMessageInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await editMessage(
    input.conversationId,
    input.messageId,
    input.content
  );

  if (!result.ok) {
    return { success: false, error: result.error };
  }

  return {
    success: true,
    data: {
      message: 'Message edited successfully',
      conversationId: result.value.conversationId,
      messageId: result.value.messageId,
    },
  };
}

async function handleDeleteMessage(
  input: z.infer<typeof DeleteMessageInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await deleteMessage(
    input.conversationId,
    input.messageId
  );

  if (!result.ok) {
    return { success: false, error: result.error };
  }

  return {
    success: true,
    data: {
      message: 'Message deleted successfully',
      conversationId: result.value.conversationId,
      messageId: result.value.messageId,
    },
  };
}

async function handleGetUnread(
  input: z.infer<typeof GetUnreadInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  // If a specific conversation is provided, just check that one
  if (input.conversationId) {
    const result = await getUnreadStatus(input.conversationId);
    if (!result.ok) {
      return { success: false, error: result.error };
    }

    return {
      success: true,
      data: {
        conversationId: result.value.conversationId,
        unreadCount: result.value.unreadCount,
        lastReadMessageId: result.value.lastReadMessageId,
        latestMessageId: result.value.latestMessageId,
      },
    };
  }

  // Aggregate mode: bulk check all conversations in a single API call
  const result = await getUnreadConversations();
  if (!result.ok) {
    return { success: false, error: result.error };
  }

  const { unreadChats, unreadChannels, totalChecked } = result.value;

  const formatConv = (c: typeof unreadChats[number]) => ({
    conversationId: c.conversationId,
    displayName: c.displayName,
    lastMessageFrom: c.lastMessageFrom,
  });

  return {
    success: true,
    data: {
      unreadChats: unreadChats.length,
      unreadChannels: unreadChannels.length,
      chats: unreadChats.map(formatConv),
      channels: unreadChannels.map(formatConv),
      totalChecked,
      note: totalChecked >= 200 ? 'Only checked most recent 200 conversations' : undefined,
    },
  };
}

async function handleMarkAsRead(
  input: z.infer<typeof MarkAsReadInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await markAsRead(input.conversationId, input.messageId);

  if (!result.ok) {
    return { success: false, error: result.error };
  }

  return {
    success: true,
    data: {
      message: 'Conversation marked as read',
      conversationId: result.value.conversationId,
      markedUpTo: result.value.markedUpTo,
    },
  };
}

async function handleGetActivity(
  input: z.infer<typeof GetActivityInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await getActivityFeed({ limit: input.limit, syncState: input.syncState });

  if (!result.ok) {
    return { success: false, error: result.error };
  }

  return {
    success: true,
    data: {
      count: result.value.activities.length,
      activities: result.value.activities,
      syncState: result.value.syncState,
    },
  };
}

async function handleSearchEmoji(
  input: z.infer<typeof SearchEmojiInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const query = input.query.toLowerCase();
  
  // Search standard emojis
  const standardMatches = STANDARD_EMOJIS.filter(emoji =>
    emoji.key.toLowerCase().includes(query) ||
    emoji.description.toLowerCase().includes(query)
  ).map(emoji => ({
    key: emoji.key,
    description: emoji.description,
    type: 'standard' as const,
    category: emoji.category,
  }));

  // Try to get custom emojis
  let customMatches: Array<{
    key: string;
    description: string;
    type: 'custom';
    shortcut: string;
  }> = [];

  const customResult = await getCustomEmojis();
  if (customResult.ok) {
    customMatches = customResult.value.emojis
      .filter(emoji =>
        emoji.shortcut.toLowerCase().includes(query) ||
        emoji.description.toLowerCase().includes(query)
      )
      .map(emoji => ({
        key: emoji.id,
        description: emoji.description,
        type: 'custom' as const,
        shortcut: emoji.shortcut,
      }));
  }

  // Combine results, standard first
  const results = [...standardMatches, ...customMatches];

  return {
    success: true,
    data: {
      query: input.query,
      count: results.length,
      emojis: results,
      note: results.length === 0
        ? 'No emojis found. Try a different search term.'
        : 'Use the "key" value with teams_add_reaction.',
    },
  };
}

async function handleAddReaction(
  input: z.infer<typeof AddReactionInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await addReaction(
    input.conversationId,
    input.messageId,
    input.emoji
  );

  if (!result.ok) {
    return { success: false, error: result.error };
  }

  return {
    success: true,
    data: {
      message: `Added ${input.emoji} reaction`,
      conversationId: result.value.conversationId,
      messageId: result.value.messageId,
      emoji: result.value.emoji,
    },
  };
}

async function handleRemoveReaction(
  input: z.infer<typeof RemoveReactionInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await removeReaction(
    input.conversationId,
    input.messageId,
    input.emoji
  );

  if (!result.ok) {
    return { success: false, error: result.error };
  }

  return {
    success: true,
    data: {
      message: `Removed ${input.emoji} reaction`,
      conversationId: result.value.conversationId,
      messageId: result.value.messageId,
      emoji: result.value.emoji,
    },
  };
}

async function handleGetSavedMessages(
  input: z.infer<typeof GetSavedMessagesInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await getSavedMessages({ limit: input.limit });

  if (!result.ok) {
    return { success: false, error: result.error };
  }

  return {
    success: true,
    data: {
      count: result.value.messages.length,
      messages: result.value.messages.map(msg => ({
        content: msg.content,
        contentType: msg.contentType,
        sender: msg.sender,
        timestamp: msg.timestamp,
        sourceConversationId: msg.sourceConversationId,
        sourceMessageId: msg.sourceMessageId,
        messageLink: msg.messageLink,
      })),
    },
  };
}

async function handleGetFollowedThreads(
  input: z.infer<typeof GetFollowedThreadsInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await getFollowedThreads({ limit: input.limit });

  if (!result.ok) {
    return { success: false, error: result.error };
  }

  return {
    success: true,
    data: {
      count: result.value.threads.length,
      threads: result.value.threads.map(thread => ({
        content: thread.content,
        contentType: thread.contentType,
        sender: thread.sender,
        timestamp: thread.timestamp,
        sourceConversationId: thread.sourceConversationId,
        sourcePostId: thread.sourcePostId,
        messageLink: thread.messageLink,
      })),
    },
  };
}

async function handleGetMessage(
  input: z.infer<typeof GetMessageInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await getMessage(input.conversationId, input.messageId);

  if (!result.ok) {
    return { success: false, error: result.error };
  }

  const msg = result.value;
  return {
    success: true,
    data: {
      id: msg.id,
      content: msg.content,
      contentType: msg.contentType,
      sender: msg.sender,
      timestamp: msg.timestamp,
      when: msg.when,
      conversationId: msg.conversationId,
      isFromMe: msg.isFromMe,
      messageLink: msg.messageLink,
      links: msg.links,
      threadRootId: msg.threadRootId,
      isThreadReply: msg.isThreadReply,
      rawHtml: msg.rawHtml,
      rawFileObjects: msg.rawFileObjects,
    },
  };
}

async function handleListChats(
  input: z.infer<typeof ListChatsInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await listConversations({ limit: input.limit, type: input.type });
  if (!result.ok) {
    return { success: false, error: result.error };
  }
  return { success: true, data: { conversations: result.value.conversations, total: result.value.total } };
}

async function handleMarkUnread(
  input: z.infer<typeof MarkUnreadInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await markUnread(input.conversationId, input.messageId);
  if (!result.ok) {
    return { success: false, error: result.error };
  }
  return { success: true, data: { ...result.value, message: 'Conversation marked unread.' } };
}

async function handleListTeams(
  _input: z.infer<typeof ListTeamsInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await getMyTeamsAndChannels();
  if (!result.ok) {
    return { success: false, error: result.error };
  }
  return { success: true, data: { count: result.value.teams.length, teams: result.value.teams } };
}

async function handleGetChatMembers(
  input: z.infer<typeof GetChatMembersInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await getChatMembers(input.conversationId);
  if (!result.ok) return { success: false, error: result.error };
  return { success: true, data: { totalMemberCount: result.value.totalMemberCount, members: result.value.members } };
}

async function handleAddMember(
  input: z.infer<typeof AddMemberInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await addMember(input.conversationId, input.userId, input.role);
  if (!result.ok) return { success: false, error: result.error };
  return { success: true, data: { ...result.value, message: 'Member added.' } };
}

async function handleRemoveMember(
  input: z.infer<typeof RemoveMemberInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await removeMember(input.conversationId, input.userId);
  if (!result.ok) return { success: false, error: result.error };
  return { success: true, data: { ...result.value, message: 'Member removed.' } };
}

async function handleLeaveChat(
  input: z.infer<typeof LeaveChatInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await leaveChat(input.conversationId);
  if (!result.ok) return { success: false, error: result.error };
  return { success: true, data: { ...result.value, message: 'Left the chat.' } };
}

async function handleRenameChat(
  input: z.infer<typeof RenameChatInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await renameChat(input.conversationId, input.topic);
  if (!result.ok) return { success: false, error: result.error };
  return { success: true, data: { ...result.value, message: 'Chat renamed.' } };
}

async function handleForwardMessage(
  input: z.infer<typeof ForwardMessageInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await forwardMessage(
    input.sourceConversationId,
    input.messageId,
    input.targetConversationId,
    input.comment
  );
  if (!result.ok) return { success: false, error: result.error };
  return { success: true, data: { ...result.value, message: 'Message forwarded.' } };
}

async function handlePinMessage(
  input: z.infer<typeof PinMessageInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await pinMessage(input.conversationId, input.messageId);
  if (!result.ok) return { success: false, error: result.error };
  return { success: true, data: { ...result.value, message: 'Message pinned.' } };
}

async function handleUnpinMessage(
  input: z.infer<typeof UnpinMessageInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await unpinMessage(input.conversationId);
  if (!result.ok) return { success: false, error: result.error };
  return { success: true, data: { ...result.value, message: 'Pinned messages cleared.' } };
}

async function handleMuteChat(
  input: z.infer<typeof MuteChatInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await setMuted(input.conversationId, true);
  if (!result.ok) return { success: false, error: result.error };
  return { success: true, data: { ...result.value, message: 'Conversation muted.' } };
}

async function handleUnmuteChat(
  input: z.infer<typeof MuteChatInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await setMuted(input.conversationId, false);
  if (!result.ok) return { success: false, error: result.error };
  return { success: true, data: { ...result.value, message: 'Conversation unmuted.' } };
}

// ─────────────────────────────────────────────────────────────────────────────
// Exports
// ─────────────────────────────────────────────────────────────────────────────

export const sendMessageTool: RegisteredTool<typeof SendMessageInputSchema> = {
  definition: sendMessageToolDefinition,
  schema: SendMessageInputSchema,
  handler: handleSendMessage,
};

export const getFavoritesTool: RegisteredTool<z.ZodObject<Record<string, never>>> = {
  definition: getFavoritesToolDefinition,
  schema: z.object({}),
  handler: handleGetFavorites,
};

export const addFavoriteTool: RegisteredTool<typeof FavoriteInputSchema> = {
  definition: addFavoriteToolDefinition,
  schema: FavoriteInputSchema,
  handler: handleAddFavorite,
};

export const removeFavoriteTool: RegisteredTool<typeof FavoriteInputSchema> = {
  definition: removeFavoriteToolDefinition,
  schema: FavoriteInputSchema,
  handler: handleRemoveFavorite,
};

export const saveMessageTool: RegisteredTool<typeof SaveMessageInputSchema> = {
  definition: saveMessageToolDefinition,
  schema: SaveMessageInputSchema,
  handler: handleSaveMessage,
};

export const unsaveMessageTool: RegisteredTool<typeof SaveMessageInputSchema> = {
  definition: unsaveMessageToolDefinition,
  schema: SaveMessageInputSchema,
  handler: handleUnsaveMessage,
};

export const getChatTool: RegisteredTool<typeof GetChatInputSchema> = {
  definition: getChatToolDefinition,
  schema: GetChatInputSchema,
  handler: handleGetChat,
};

export const createGroupChatTool: RegisteredTool<typeof CreateGroupChatInputSchema> = {
  definition: createGroupChatToolDefinition,
  schema: CreateGroupChatInputSchema,
  handler: handleCreateGroupChat,
};

export const editMessageTool: RegisteredTool<typeof EditMessageInputSchema> = {
  definition: editMessageToolDefinition,
  schema: EditMessageInputSchema,
  handler: handleEditMessage,
};

export const deleteMessageTool: RegisteredTool<typeof DeleteMessageInputSchema> = {
  definition: deleteMessageToolDefinition,
  schema: DeleteMessageInputSchema,
  handler: handleDeleteMessage,
};

export const getUnreadTool: RegisteredTool<typeof GetUnreadInputSchema> = {
  definition: getUnreadToolDefinition,
  schema: GetUnreadInputSchema,
  handler: handleGetUnread,
};

export const markAsReadTool: RegisteredTool<typeof MarkAsReadInputSchema> = {
  definition: markAsReadToolDefinition,
  schema: MarkAsReadInputSchema,
  handler: handleMarkAsRead,
};

export const getActivityTool: RegisteredTool<typeof GetActivityInputSchema> = {
  definition: getActivityToolDefinition,
  schema: GetActivityInputSchema,
  handler: handleGetActivity,
};

export const searchEmojiTool: RegisteredTool<typeof SearchEmojiInputSchema> = {
  definition: searchEmojiToolDefinition,
  schema: SearchEmojiInputSchema,
  handler: handleSearchEmoji,
};

export const addReactionTool: RegisteredTool<typeof AddReactionInputSchema> = {
  definition: addReactionToolDefinition,
  schema: AddReactionInputSchema,
  handler: handleAddReaction,
};

export const removeReactionTool: RegisteredTool<typeof RemoveReactionInputSchema> = {
  definition: removeReactionToolDefinition,
  schema: RemoveReactionInputSchema,
  handler: handleRemoveReaction,
};

export const getSavedMessagesTool: RegisteredTool<typeof GetSavedMessagesInputSchema> = {
  definition: getSavedMessagesToolDefinition,
  schema: GetSavedMessagesInputSchema,
  handler: handleGetSavedMessages,
};

export const getFollowedThreadsTool: RegisteredTool<typeof GetFollowedThreadsInputSchema> = {
  definition: getFollowedThreadsToolDefinition,
  schema: GetFollowedThreadsInputSchema,
  handler: handleGetFollowedThreads,
};

export const getMessageTool: RegisteredTool<typeof GetMessageInputSchema> = {
  definition: getMessageToolDefinition,
  schema: GetMessageInputSchema,
  handler: handleGetMessage,
};

export const listChatsTool: RegisteredTool<typeof ListChatsInputSchema> = {
  definition: listChatsToolDefinition,
  schema: ListChatsInputSchema,
  handler: handleListChats,
};

export const markUnreadTool: RegisteredTool<typeof MarkUnreadInputSchema> = {
  definition: markUnreadToolDefinition,
  schema: MarkUnreadInputSchema,
  handler: handleMarkUnread,
};

export const listTeamsTool: RegisteredTool<typeof ListTeamsInputSchema> = {
  definition: listTeamsToolDefinition,
  schema: ListTeamsInputSchema,
  handler: handleListTeams,
};

export const getChatMembersTool: RegisteredTool<typeof GetChatMembersInputSchema> = {
  definition: getChatMembersToolDefinition,
  schema: GetChatMembersInputSchema,
  handler: handleGetChatMembers,
};

export const addMemberTool: RegisteredTool<typeof AddMemberInputSchema> = {
  definition: addMemberToolDefinition,
  schema: AddMemberInputSchema,
  handler: handleAddMember,
};

export const removeMemberTool: RegisteredTool<typeof RemoveMemberInputSchema> = {
  definition: removeMemberToolDefinition,
  schema: RemoveMemberInputSchema,
  handler: handleRemoveMember,
};

export const leaveChatTool: RegisteredTool<typeof LeaveChatInputSchema> = {
  definition: leaveChatToolDefinition,
  schema: LeaveChatInputSchema,
  handler: handleLeaveChat,
};

export const renameChatTool: RegisteredTool<typeof RenameChatInputSchema> = {
  definition: renameChatToolDefinition,
  schema: RenameChatInputSchema,
  handler: handleRenameChat,
};

export const forwardMessageTool: RegisteredTool<typeof ForwardMessageInputSchema> = {
  definition: forwardMessageToolDefinition,
  schema: ForwardMessageInputSchema,
  handler: handleForwardMessage,
};

export const pinMessageTool: RegisteredTool<typeof PinMessageInputSchema> = {
  definition: pinMessageToolDefinition, schema: PinMessageInputSchema, handler: handlePinMessage,
};
export const unpinMessageTool: RegisteredTool<typeof UnpinMessageInputSchema> = {
  definition: unpinMessageToolDefinition, schema: UnpinMessageInputSchema, handler: handleUnpinMessage,
};
export const muteChatTool: RegisteredTool<typeof MuteChatInputSchema> = {
  definition: muteChatToolDefinition, schema: MuteChatInputSchema, handler: handleMuteChat,
};
export const unmuteChatTool: RegisteredTool<typeof MuteChatInputSchema> = {
  definition: unmuteChatToolDefinition, schema: MuteChatInputSchema, handler: handleUnmuteChat,
};

/** All message-related tools. */
export const messageTools = [
  sendMessageTool,
  getFavoritesTool,
  addFavoriteTool,
  removeFavoriteTool,
  saveMessageTool,
  unsaveMessageTool,
  getChatTool,
  createGroupChatTool,
  editMessageTool,
  deleteMessageTool,
  getUnreadTool,
  markAsReadTool,
  getActivityTool,
  searchEmojiTool,
  addReactionTool,
  removeReactionTool,
  getSavedMessagesTool,
  getFollowedThreadsTool,
  getMessageTool,
  listChatsTool,
  markUnreadTool,
  listTeamsTool,
  getChatMembersTool,
  addMemberTool,
  removeMemberTool,
  leaveChatTool,
  renameChatTool,
  forwardMessageTool,
  pinMessageTool,
  unpinMessageTool,
  muteChatTool,
  unmuteChatTool,
];
