/**
 * Unit tests for message tools (real handlers).
 */
import { describe, it, expect, vi, beforeEach } from 'vitest';
import { ok, err } from '../types/result.js';
import { createError, ErrorCode } from '../types/errors.js';
import './index.js';

vi.mock('../api/chatsvc-api.js', () => ({
  sendMessage: vi.fn(), getMessage: vi.fn(), saveMessage: vi.fn(), unsaveMessage: vi.fn(),
  getOneOnOneChatId: vi.fn(), createGroupChat: vi.fn(), editMessage: vi.fn(), deleteMessage: vi.fn(),
  getUnreadStatus: vi.fn(), getUnreadConversations: vi.fn(), markAsRead: vi.fn(), getActivityFeed: vi.fn(),
  addReaction: vi.fn(), removeReaction: vi.fn(), getSavedMessages: vi.fn(), getFollowedThreads: vi.fn(),
  listConversations: vi.fn(), markUnread: vi.fn(), getChatMembers: vi.fn(), addMember: vi.fn(),
  removeMember: vi.fn(), leaveChat: vi.fn(), renameChat: vi.fn(), forwardMessage: vi.fn(),
  pinMessage: vi.fn(), unpinMessage: vi.fn(), setMuted: vi.fn(),
}));
vi.mock('../api/csa-api.js', () => ({
  getFavorites: vi.fn(), addFavorite: vi.fn(), removeFavorite: vi.fn(),
  getCustomEmojis: vi.fn(), getMyTeamsAndChannels: vi.fn(),
}));

import * as chat from '../api/chatsvc-api.js';
import * as csa from '../api/csa-api.js';
import * as M from './message-tools.js';

const ctx = { server: {} } as never;
const anErr = err(createError(ErrorCode.API_ERROR, 'boom'));

beforeEach(() => vi.clearAllMocks());

describe('sendMessageTool', () => {
  it('defaults conversationId to self-chat', () => {
    expect(M.SendMessageInputSchema.parse({ content: 'hi' }).conversationId).toBe('48:notes');
  });

  it('returns thread note when replying', async () => {
    vi.mocked(chat.sendMessage).mockResolvedValue(ok({ messageId: 'm1', timestamp: 1700 }) as never);
    const res = await M.sendMessageTool.handler(
      { content: 'hi', conversationId: 'c1', replyToMessageId: 'r1' }, ctx
    );
    expect(res.success).toBe(true);
    if (res.success) {
      expect(res.data.serverMessageId).toBe('1700');
      expect(String(res.data.note)).toContain('reply');
    }
  });

  it('returns serverMessageId note without reply', async () => {
    vi.mocked(chat.sendMessage).mockResolvedValue(ok({ messageId: 'm1', timestamp: 1700 }) as never);
    const res = await M.sendMessageTool.handler({ content: 'hi', conversationId: 'c1' }, ctx);
    expect(res.success).toBe(true);
    if (res.success) expect(String(res.data.note)).toContain('Use serverMessageId');
  });

  it('handles missing timestamp (no serverMessageId)', async () => {
    vi.mocked(chat.sendMessage).mockResolvedValue(ok({ messageId: 'm1', timestamp: undefined }) as never);
    const res = await M.sendMessageTool.handler({ content: 'hi', conversationId: 'c1' }, ctx);
    expect(res.success).toBe(true);
    if (res.success) {
      expect(res.data.serverMessageId).toBeUndefined();
      expect(res.data.note).toBeUndefined();
    }
  });

  it('propagates errors', async () => {
    vi.mocked(chat.sendMessage).mockResolvedValue(anErr as never);
    expect((await M.sendMessageTool.handler({ content: 'hi', conversationId: 'c1' }, ctx)).success).toBe(false);
  });
});

describe('favorites', () => {
  it('getFavorites success', async () => {
    vi.mocked(csa.getFavorites).mockResolvedValue(ok({ favorites: [{ id: 'f1' }] }) as never);
    const res = await M.getFavoritesTool.handler({}, ctx);
    expect(res.success).toBe(true);
    if (res.success) expect(res.data.count).toBe(1);
  });
  it('getFavorites error', async () => {
    vi.mocked(csa.getFavorites).mockResolvedValue(anErr as never);
    expect((await M.getFavoritesTool.handler({}, ctx)).success).toBe(false);
  });
  it('addFavorite success', async () => {
    vi.mocked(csa.addFavorite).mockResolvedValue(ok({}) as never);
    expect((await M.addFavoriteTool.handler({ conversationId: 'c1' }, ctx)).success).toBe(true);
  });
  it('addFavorite error', async () => {
    vi.mocked(csa.addFavorite).mockResolvedValue(anErr as never);
    expect((await M.addFavoriteTool.handler({ conversationId: 'c1' }, ctx)).success).toBe(false);
  });
  it('removeFavorite success', async () => {
    vi.mocked(csa.removeFavorite).mockResolvedValue(ok({}) as never);
    expect((await M.removeFavoriteTool.handler({ conversationId: 'c1' }, ctx)).success).toBe(true);
  });
  it('removeFavorite error', async () => {
    vi.mocked(csa.removeFavorite).mockResolvedValue(anErr as never);
    expect((await M.removeFavoriteTool.handler({ conversationId: 'c1' }, ctx)).success).toBe(false);
  });
});

describe('save / unsave', () => {
  it('save success', async () => {
    vi.mocked(chat.saveMessage).mockResolvedValue(ok({}) as never);
    expect((await M.saveMessageTool.handler({ conversationId: 'c1', messageId: 'm1' }, ctx)).success).toBe(true);
  });
  it('save error', async () => {
    vi.mocked(chat.saveMessage).mockResolvedValue(anErr as never);
    expect((await M.saveMessageTool.handler({ conversationId: 'c1', messageId: 'm1' }, ctx)).success).toBe(false);
  });
  it('unsave success', async () => {
    vi.mocked(chat.unsaveMessage).mockResolvedValue(ok({}) as never);
    expect((await M.unsaveMessageTool.handler({ conversationId: 'c1', messageId: 'm1' }, ctx)).success).toBe(true);
  });
  it('unsave error', async () => {
    vi.mocked(chat.unsaveMessage).mockResolvedValue(anErr as never);
    expect((await M.unsaveMessageTool.handler({ conversationId: 'c1', messageId: 'm1' }, ctx)).success).toBe(false);
  });
});

describe('getChat', () => {
  it('success (sync result)', async () => {
    vi.mocked(chat.getOneOnOneChatId).mockReturnValue(
      ok({ conversationId: 'c1', otherUserId: 'o', currentUserId: 'u' }) as never
    );
    const res = await M.getChatTool.handler({ userId: 'u1' }, ctx);
    expect(res.success).toBe(true);
    if (res.success) expect(res.data.conversationId).toBe('c1');
  });
  it('error', async () => {
    vi.mocked(chat.getOneOnOneChatId).mockReturnValue(anErr as never);
    expect((await M.getChatTool.handler({ userId: 'u1' }, ctx)).success).toBe(false);
  });
});

describe('createGroupChat', () => {
  it('success', async () => {
    vi.mocked(chat.createGroupChat).mockResolvedValue(
      ok({ conversationId: 'c1', members: [], topic: 't' }) as never
    );
    const res = await M.createGroupChatTool.handler({ userIds: ['a', 'b'] }, ctx);
    expect(res.success).toBe(true);
  });
  it('error', async () => {
    vi.mocked(chat.createGroupChat).mockResolvedValue(anErr as never);
    expect((await M.createGroupChatTool.handler({ userIds: ['a', 'b'] }, ctx)).success).toBe(false);
  });
});

describe('edit / delete', () => {
  it('edit success', async () => {
    vi.mocked(chat.editMessage).mockResolvedValue(ok({ conversationId: 'c1', messageId: 'm1' }) as never);
    expect((await M.editMessageTool.handler({ conversationId: 'c1', messageId: 'm1', content: 'x' }, ctx)).success).toBe(true);
  });
  it('edit error', async () => {
    vi.mocked(chat.editMessage).mockResolvedValue(anErr as never);
    expect((await M.editMessageTool.handler({ conversationId: 'c1', messageId: 'm1', content: 'x' }, ctx)).success).toBe(false);
  });
  it('delete success', async () => {
    vi.mocked(chat.deleteMessage).mockResolvedValue(ok({ conversationId: 'c1', messageId: 'm1' }) as never);
    expect((await M.deleteMessageTool.handler({ conversationId: 'c1', messageId: 'm1' }, ctx)).success).toBe(true);
  });
  it('delete error', async () => {
    vi.mocked(chat.deleteMessage).mockResolvedValue(anErr as never);
    expect((await M.deleteMessageTool.handler({ conversationId: 'c1', messageId: 'm1' }, ctx)).success).toBe(false);
  });
});

describe('getUnread', () => {
  it('single conversation success', async () => {
    vi.mocked(chat.getUnreadStatus).mockResolvedValue(
      ok({ conversationId: 'c1', unreadCount: 2, lastReadMessageId: 'm1', latestMessageId: 'm3' }) as never
    );
    const res = await M.getUnreadTool.handler({ conversationId: 'c1' }, ctx);
    expect(res.success).toBe(true);
    if (res.success) expect(res.data.unreadCount).toBe(2);
  });
  it('single conversation error', async () => {
    vi.mocked(chat.getUnreadStatus).mockResolvedValue(anErr as never);
    expect((await M.getUnreadTool.handler({ conversationId: 'c1' }, ctx)).success).toBe(false);
  });
  it('bulk success with note when at cap', async () => {
    vi.mocked(chat.getUnreadConversations).mockResolvedValue(ok({
      unreadChats: [{ conversationId: 'c1', displayName: 'A', lastMessageFrom: 'x' }],
      unreadChannels: [],
      totalChecked: 200,
    }) as never);
    const res = await M.getUnreadTool.handler({}, ctx);
    expect(res.success).toBe(true);
    if (res.success) {
      expect(res.data.unreadChats).toBe(1);
      expect(String(res.data.note)).toContain('200');
    }
  });
  it('bulk success without note', async () => {
    vi.mocked(chat.getUnreadConversations).mockResolvedValue(ok({
      unreadChats: [], unreadChannels: [], totalChecked: 5,
    }) as never);
    const res = await M.getUnreadTool.handler({}, ctx);
    expect(res.success).toBe(true);
    if (res.success) expect(res.data.note).toBeUndefined();
  });
  it('bulk error', async () => {
    vi.mocked(chat.getUnreadConversations).mockResolvedValue(anErr as never);
    expect((await M.getUnreadTool.handler({}, ctx)).success).toBe(false);
  });
});

describe('markAsRead', () => {
  it('success', async () => {
    vi.mocked(chat.markAsRead).mockResolvedValue(ok({ conversationId: 'c1', markedUpTo: 'm1' }) as never);
    expect((await M.markAsReadTool.handler({ conversationId: 'c1', messageId: 'm1' }, ctx)).success).toBe(true);
  });
  it('error', async () => {
    vi.mocked(chat.markAsRead).mockResolvedValue(anErr as never);
    expect((await M.markAsReadTool.handler({ conversationId: 'c1', messageId: 'm1' }, ctx)).success).toBe(false);
  });
});

describe('getActivity', () => {
  it('success', async () => {
    vi.mocked(chat.getActivityFeed).mockResolvedValue(ok({ activities: [{}], syncState: 's' }) as never);
    const res = await M.getActivityTool.handler({}, ctx);
    expect(res.success).toBe(true);
    if (res.success) expect(res.data.count).toBe(1);
  });
  it('error', async () => {
    vi.mocked(chat.getActivityFeed).mockResolvedValue(anErr as never);
    expect((await M.getActivityTool.handler({}, ctx)).success).toBe(false);
  });
});

describe('searchEmoji', () => {
  it('returns standard and custom matches', async () => {
    vi.mocked(csa.getCustomEmojis).mockResolvedValue(ok({
      emojis: [{ id: 'e1', shortcut: 'likeparty', description: 'party like' }],
    }) as never);
    const res = await M.searchEmojiTool.handler({ query: 'like' }, ctx);
    expect(res.success).toBe(true);
    if (res.success) expect((res.data.count as number)).toBeGreaterThan(0);
  });
  it('handles no matches (note) and failed custom lookup', async () => {
    vi.mocked(csa.getCustomEmojis).mockResolvedValue(anErr as never);
    const res = await M.searchEmojiTool.handler({ query: 'zzzznotanemoji' }, ctx);
    expect(res.success).toBe(true);
    if (res.success) {
      expect(res.data.count).toBe(0);
      expect(String(res.data.note)).toContain('No emojis found');
    }
  });
});

describe('reactions', () => {
  it('add success', async () => {
    vi.mocked(chat.addReaction).mockResolvedValue(ok({ conversationId: 'c1', messageId: 'm1', emoji: 'like' }) as never);
    expect((await M.addReactionTool.handler({ conversationId: 'c1', messageId: 'm1', emoji: 'like' }, ctx)).success).toBe(true);
  });
  it('add error', async () => {
    vi.mocked(chat.addReaction).mockResolvedValue(anErr as never);
    expect((await M.addReactionTool.handler({ conversationId: 'c1', messageId: 'm1', emoji: 'like' }, ctx)).success).toBe(false);
  });
  it('remove success', async () => {
    vi.mocked(chat.removeReaction).mockResolvedValue(ok({ conversationId: 'c1', messageId: 'm1', emoji: 'like' }) as never);
    expect((await M.removeReactionTool.handler({ conversationId: 'c1', messageId: 'm1', emoji: 'like' }, ctx)).success).toBe(true);
  });
  it('remove error', async () => {
    vi.mocked(chat.removeReaction).mockResolvedValue(anErr as never);
    expect((await M.removeReactionTool.handler({ conversationId: 'c1', messageId: 'm1', emoji: 'like' }, ctx)).success).toBe(false);
  });
});

describe('saved messages / followed threads', () => {
  it('saved success', async () => {
    vi.mocked(chat.getSavedMessages).mockResolvedValue(ok({ messages: [{
      content: 'c', contentType: 't', sender: 's', timestamp: 'ts',
      sourceConversationId: 'sc', sourceMessageId: 'sm', messageLink: 'ml',
    }] }) as never);
    const res = await M.getSavedMessagesTool.handler({}, ctx);
    expect(res.success).toBe(true);
    if (res.success) expect(res.data.count).toBe(1);
  });
  it('saved error', async () => {
    vi.mocked(chat.getSavedMessages).mockResolvedValue(anErr as never);
    expect((await M.getSavedMessagesTool.handler({}, ctx)).success).toBe(false);
  });
  it('followed success', async () => {
    vi.mocked(chat.getFollowedThreads).mockResolvedValue(ok({ threads: [{
      content: 'c', contentType: 't', sender: 's', timestamp: 'ts',
      sourceConversationId: 'sc', sourcePostId: 'sp', messageLink: 'ml',
    }] }) as never);
    const res = await M.getFollowedThreadsTool.handler({}, ctx);
    expect(res.success).toBe(true);
    if (res.success) expect(res.data.count).toBe(1);
  });
  it('followed error', async () => {
    vi.mocked(chat.getFollowedThreads).mockResolvedValue(anErr as never);
    expect((await M.getFollowedThreadsTool.handler({}, ctx)).success).toBe(false);
  });
});

describe('getMessage', () => {
  it('success', async () => {
    vi.mocked(chat.getMessage).mockResolvedValue(ok({
      id: 'm1', content: 'c', contentType: 't', sender: 's', timestamp: 'ts', when: 'w',
      conversationId: 'c1', isFromMe: false, messageLink: 'ml', links: [], threadRootId: undefined, isThreadReply: false,
    }) as never);
    const res = await M.getMessageTool.handler({ conversationId: 'c1', messageId: 'm1' }, ctx);
    expect(res.success).toBe(true);
    if (res.success) expect(res.data.id).toBe('m1');
  });
  it('error', async () => {
    vi.mocked(chat.getMessage).mockResolvedValue(anErr as never);
    expect((await M.getMessageTool.handler({ conversationId: 'c1', messageId: 'm1' }, ctx)).success).toBe(false);
  });
});

describe('listChats', () => {
  it('defaults limit to 50', () => {
    expect(M.ListChatsInputSchema.parse({}).limit).toBe(50);
  });
  it('success', async () => {
    vi.mocked(chat.listConversations).mockResolvedValue(ok({ conversations: [], total: 0 }) as never);
    expect((await M.listChatsTool.handler({ limit: 50 }, ctx)).success).toBe(true);
  });
  it('error', async () => {
    vi.mocked(chat.listConversations).mockResolvedValue(anErr as never);
    expect((await M.listChatsTool.handler({ limit: 50 }, ctx)).success).toBe(false);
  });
});

describe('markUnread', () => {
  it('success', async () => {
    vi.mocked(chat.markUnread).mockResolvedValue(ok({ conversationId: 'c1' }) as never);
    expect((await M.markUnreadTool.handler({ conversationId: 'c1', messageId: 'm1' }, ctx)).success).toBe(true);
  });
  it('error', async () => {
    vi.mocked(chat.markUnread).mockResolvedValue(anErr as never);
    expect((await M.markUnreadTool.handler({ conversationId: 'c1', messageId: 'm1' }, ctx)).success).toBe(false);
  });
});

describe('listTeams', () => {
  it('success', async () => {
    vi.mocked(csa.getMyTeamsAndChannels).mockResolvedValue(ok({ teams: [{ id: 't1' }] }) as never);
    const res = await M.listTeamsTool.handler({}, ctx);
    expect(res.success).toBe(true);
    if (res.success) expect(res.data.count).toBe(1);
  });
  it('error', async () => {
    vi.mocked(csa.getMyTeamsAndChannels).mockResolvedValue(anErr as never);
    expect((await M.listTeamsTool.handler({}, ctx)).success).toBe(false);
  });
});

describe('chat members and membership', () => {
  it('getChatMembers success', async () => {
    vi.mocked(chat.getChatMembers).mockResolvedValue(ok({ totalMemberCount: 2, members: [] }) as never);
    expect((await M.getChatMembersTool.handler({ conversationId: 'c1' }, ctx)).success).toBe(true);
  });
  it('getChatMembers error', async () => {
    vi.mocked(chat.getChatMembers).mockResolvedValue(anErr as never);
    expect((await M.getChatMembersTool.handler({ conversationId: 'c1' }, ctx)).success).toBe(false);
  });
  it('addMember default role and success', async () => {
    expect(M.AddMemberInputSchema.parse({ conversationId: 'c1', userId: 'u1' }).role).toBe('User');
    vi.mocked(chat.addMember).mockResolvedValue(ok({}) as never);
    expect((await M.addMemberTool.handler({ conversationId: 'c1', userId: 'u1', role: 'User' }, ctx)).success).toBe(true);
  });
  it('addMember error', async () => {
    vi.mocked(chat.addMember).mockResolvedValue(anErr as never);
    expect((await M.addMemberTool.handler({ conversationId: 'c1', userId: 'u1', role: 'User' }, ctx)).success).toBe(false);
  });
  it('removeMember success/error', async () => {
    vi.mocked(chat.removeMember).mockResolvedValueOnce(ok({}) as never).mockResolvedValueOnce(anErr as never);
    expect((await M.removeMemberTool.handler({ conversationId: 'c1', userId: 'u1' }, ctx)).success).toBe(true);
    expect((await M.removeMemberTool.handler({ conversationId: 'c1', userId: 'u1' }, ctx)).success).toBe(false);
  });
  it('leaveChat success/error', async () => {
    vi.mocked(chat.leaveChat).mockResolvedValueOnce(ok({}) as never).mockResolvedValueOnce(anErr as never);
    expect((await M.leaveChatTool.handler({ conversationId: 'c1' }, ctx)).success).toBe(true);
    expect((await M.leaveChatTool.handler({ conversationId: 'c1' }, ctx)).success).toBe(false);
  });
  it('renameChat success/error', async () => {
    vi.mocked(chat.renameChat).mockResolvedValueOnce(ok({}) as never).mockResolvedValueOnce(anErr as never);
    expect((await M.renameChatTool.handler({ conversationId: 'c1', topic: 't' }, ctx)).success).toBe(true);
    expect((await M.renameChatTool.handler({ conversationId: 'c1', topic: 't' }, ctx)).success).toBe(false);
  });
});

describe('forward / pin / unpin / mute / unmute', () => {
  it('forward success/error', async () => {
    vi.mocked(chat.forwardMessage).mockResolvedValueOnce(ok({}) as never).mockResolvedValueOnce(anErr as never);
    const input = { sourceConversationId: 's', messageId: 'm', targetConversationId: 't' };
    expect((await M.forwardMessageTool.handler(input, ctx)).success).toBe(true);
    expect((await M.forwardMessageTool.handler(input, ctx)).success).toBe(false);
  });
  it('pin success/error', async () => {
    vi.mocked(chat.pinMessage).mockResolvedValueOnce(ok({}) as never).mockResolvedValueOnce(anErr as never);
    expect((await M.pinMessageTool.handler({ conversationId: 'c1', messageId: 'm1' }, ctx)).success).toBe(true);
    expect((await M.pinMessageTool.handler({ conversationId: 'c1', messageId: 'm1' }, ctx)).success).toBe(false);
  });
  it('unpin success/error', async () => {
    vi.mocked(chat.unpinMessage).mockResolvedValueOnce(ok({}) as never).mockResolvedValueOnce(anErr as never);
    expect((await M.unpinMessageTool.handler({ conversationId: 'c1' }, ctx)).success).toBe(true);
    expect((await M.unpinMessageTool.handler({ conversationId: 'c1' }, ctx)).success).toBe(false);
  });
  it('mute success/error', async () => {
    vi.mocked(chat.setMuted).mockResolvedValueOnce(ok({}) as never).mockResolvedValueOnce(anErr as never);
    expect((await M.muteChatTool.handler({ conversationId: 'c1' }, ctx)).success).toBe(true);
    expect((await M.muteChatTool.handler({ conversationId: 'c1' }, ctx)).success).toBe(false);
  });
  it('unmute success/error', async () => {
    vi.mocked(chat.setMuted).mockResolvedValueOnce(ok({}) as never).mockResolvedValueOnce(anErr as never);
    expect((await M.unmuteChatTool.handler({ conversationId: 'c1' }, ctx)).success).toBe(true);
    expect((await M.unmuteChatTool.handler({ conversationId: 'c1' }, ctx)).success).toBe(false);
  });
});
