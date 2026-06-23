/**
 * Chat Service API - Conversation management.
 *
 * Members (list / add / remove / leave), rename, and message forwarding for
 * group chats and threads.
 */

import { tmpdir } from 'node:os';
import { unlink } from 'node:fs/promises';
import { join } from 'node:path';
import { randomBytes } from 'node:crypto';
import { httpRequest } from '../utils/http.js';
import { CHATSVC_API, getMessagingHeaders, getSkypeAuthHeaders } from '../utils/api-config.js';
import { type Result, ok, err } from '../types/result.js';
import { ErrorCode, createError } from '../types/errors.js';
import { requireMessageAuthWithConfig } from '../utils/auth-guards.js';
import { extractObjectId, escapeHtmlChars } from '../utils/parsers.js';
import { MRI_ORGID_PREFIX } from '../constants.js';
import { getMessage, sendMessage, buildReplyQuoteHtml } from './chatsvc-messaging.js';
import { downloadSharedFile, prepareFileForChat } from './files-graph-api.js';

// ─────────────────────────────────────────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────────────────────────────────────────

/** Normalise a user identifier (MRI, guid@tenant, or raw guid) to an MRI. */
function toMemberMri(identifier: string): string | null {
  if (identifier.startsWith(MRI_ORGID_PREFIX)) return identifier;
  const objectId = extractObjectId(identifier);
  return objectId ? `${MRI_ORGID_PREFIX}${objectId}` : null;
}

// ─────────────────────────────────────────────────────────────────────────────
// Members
// ─────────────────────────────────────────────────────────────────────────────

export interface ChatMember {
  mri: string;
  role: string;
  isCurrentUser: boolean;
}

export interface ChatMembersResult {
  conversationId: string;
  totalMemberCount: number;
  members: ChatMember[];
}

/** Lists the members of a group chat or channel thread. */
export async function getChatMembers(conversationId: string): Promise<Result<ChatMembersResult>> {
  const authResult = requireMessageAuthWithConfig();
  if (!authResult.ok) return authResult;
  const { auth, region, baseUrl } = authResult.value;

  const response = await httpRequest<Record<string, unknown>>(
    CHATSVC_API.threadMembers(region, conversationId, baseUrl),
    { method: 'GET', headers: getSkypeAuthHeaders(auth.skypeToken, auth.authToken, baseUrl) }
  );
  if (!response.ok) return response;

  const data = response.value.data;
  const rawMembers = (data.members as Array<Record<string, unknown>>) ?? [];
  const members: ChatMember[] = rawMembers.map((m) => ({
    mri: m.id as string,
    role: (m.role as string) ?? 'User',
    isCurrentUser: m.id === auth.userMri,
  }));

  return ok({
    conversationId,
    totalMemberCount: (data.totalMemberCount as number) ?? members.length,
    members,
  });
}

/** Adds a member to a group chat. */
export async function addMember(
  conversationId: string,
  userIdentifier: string,
  role: 'Admin' | 'User' = 'User'
): Promise<Result<{ conversationId: string; addedMri: string }>> {
  const authResult = requireMessageAuthWithConfig();
  if (!authResult.ok) return authResult;
  const { auth, region, baseUrl } = authResult.value;

  const mri = toMemberMri(userIdentifier);
  if (!mri) {
    return err(createError(ErrorCode.INVALID_INPUT, `Invalid user identifier: ${userIdentifier}.`));
  }

  const response = await httpRequest<unknown>(
    CHATSVC_API.threadMember(region, conversationId, mri, baseUrl),
    {
      method: 'PUT',
      headers: getMessagingHeaders(auth.skypeToken, auth.authToken, baseUrl),
      body: JSON.stringify({ role }),
    }
  );
  if (!response.ok) return response;
  return ok({ conversationId, addedMri: mri });
}

/** Removes a member from a group chat. */
export async function removeMember(
  conversationId: string,
  userIdentifier: string
): Promise<Result<{ conversationId: string; removedMri: string }>> {
  const authResult = requireMessageAuthWithConfig();
  if (!authResult.ok) return authResult;
  const { auth, region, baseUrl } = authResult.value;

  const mri = toMemberMri(userIdentifier);
  if (!mri) {
    return err(createError(ErrorCode.INVALID_INPUT, `Invalid user identifier: ${userIdentifier}.`));
  }

  const response = await httpRequest<unknown>(
    CHATSVC_API.threadMember(region, conversationId, mri, baseUrl),
    { method: 'DELETE', headers: getMessagingHeaders(auth.skypeToken, auth.authToken, baseUrl) }
  );
  if (!response.ok) return response;
  return ok({ conversationId, removedMri: mri });
}

/** Leaves a group chat by removing the current user. */
export async function leaveChat(conversationId: string): Promise<Result<{ conversationId: string }>> {
  const authResult = requireMessageAuthWithConfig();
  if (!authResult.ok) return authResult;
  const { auth, region, baseUrl } = authResult.value;

  const response = await httpRequest<unknown>(
    CHATSVC_API.threadMember(region, conversationId, auth.userMri, baseUrl),
    { method: 'DELETE', headers: getMessagingHeaders(auth.skypeToken, auth.authToken, baseUrl) }
  );
  if (!response.ok) return response;
  return ok({ conversationId });
}

// ─────────────────────────────────────────────────────────────────────────────
// Rename
// ─────────────────────────────────────────────────────────────────────────────

/** Renames a group chat (sets its topic). */
export async function renameChat(
  conversationId: string,
  topic: string
): Promise<Result<{ conversationId: string; topic: string }>> {
  const authResult = requireMessageAuthWithConfig();
  if (!authResult.ok) return authResult;
  const { auth, region, baseUrl } = authResult.value;

  const response = await httpRequest<unknown>(
    CHATSVC_API.threadProperty(region, conversationId, 'topic', baseUrl),
    {
      method: 'PUT',
      headers: getMessagingHeaders(auth.skypeToken, auth.authToken, baseUrl),
      body: JSON.stringify({ topic }),
    }
  );
  if (!response.ok) return response;
  return ok({ conversationId, topic });
}

// ─────────────────────────────────────────────────────────────────────────────
// Forward
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Forwards a message to another conversation. Teams has no native forward, so
 * the original is fetched and re-sent as a quoted block (with an optional note)
 * into the target conversation. If the original message has file attachments,
 * each file is downloaded from SharePoint and re-uploaded as a native chiclet
 * into the target conversation on the same message.
 */
export async function forwardMessage(
  sourceConversationId: string,
  messageId: string,
  targetConversationId: string,
  comment?: string
): Promise<Result<{ targetConversationId: string }>> {
  const original = await getMessage(sourceConversationId, messageId);
  if (!original.ok) return original;

  const quote = buildReplyQuoteHtml(original.value);
  const note = comment ? `${escapeHtmlChars(comment)}<br><br>` : '';
  const html = `${note}<i>Forwarded message:</i><br>${quote}`;

  // If the original message has file attachments, download and re-upload each one.
  const fileUrls = original.value.fileUrls ?? [];
  const fileProperties: Record<string, unknown>[] = [];
  const tempPaths: string[] = [];

  for (const url of fileUrls) {
    // Use a random temp path with the original file extension where possible.
    const ext = url.split('.').pop()?.split('?')[0] ?? 'bin';
    const tempPath = join(tmpdir(), `mcp-fwd-${randomBytes(8).toString('hex')}.${ext}`);
    tempPaths.push(tempPath);

    const downloaded = await downloadSharedFile(url, tempPath);
    if (!downloaded.ok) continue; // Skip files we cannot access; still forward the text.

    const prepared = await prepareFileForChat(targetConversationId, tempPath);
    if (prepared.ok) {
      fileProperties.push(prepared.value.fileProperty);
    }
  }

  const sent = await sendMessage(targetConversationId, '', {
    rawContentHtml: html,
    ...(fileProperties.length > 0 ? { files: fileProperties } : {}),
  });

  // Clean up temp files regardless of send outcome.
  await Promise.allSettled(tempPaths.map((p) => unlink(p)));

  if (!sent.ok) return sent;
  return ok({ targetConversationId });
}

// ─────────────────────────────────────────────────────────────────────────────
// Pin / unpin messages
// ─────────────────────────────────────────────────────────────────────────────

interface PinnedItem { itemId: string; itemType: string }

async function setPinnedItems(
  conversationId: string,
  items: PinnedItem[]
): Promise<Result<unknown>> {
  const authResult = requireMessageAuthWithConfig();
  if (!authResult.ok) return authResult;
  const { auth, region, baseUrl } = authResult.value;

  return httpRequest<unknown>(
    CHATSVC_API.threadProperty(region, conversationId, 'pinnedItems', baseUrl),
    {
      method: 'PUT',
      headers: getMessagingHeaders(auth.skypeToken, auth.authToken, baseUrl),
      // pinnedItems is a JSON-encoded string of the array, not a raw array.
      body: JSON.stringify({ pinnedItems: JSON.stringify(items) }),
    }
  );
}

/** Pins a message in a conversation (it appears in the chat's pinned items). */
export async function pinMessage(
  conversationId: string,
  messageId: string
): Promise<Result<{ conversationId: string; messageId: string }>> {
  const response = await setPinnedItems(conversationId, [{ itemId: messageId, itemType: 'Message' }]);
  if (!response.ok) return response;
  return ok({ conversationId, messageId });
}

/** Removes all pinned messages from a conversation. */
export async function unpinMessage(
  conversationId: string
): Promise<Result<{ conversationId: string }>> {
  const response = await setPinnedItems(conversationId, []);
  if (!response.ok) return response;
  return ok({ conversationId });
}

// ─────────────────────────────────────────────────────────────────────────────
// Mute / unmute
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Mutes (alerts off) or unmutes (alerts on) a conversation for the current user.
 */
export async function setMuted(
  conversationId: string,
  muted: boolean
): Promise<Result<{ conversationId: string; muted: boolean }>> {
  const authResult = requireMessageAuthWithConfig();
  if (!authResult.ok) return authResult;
  const { auth, region, baseUrl } = authResult.value;

  const response = await httpRequest<unknown>(
    CHATSVC_API.conversationProperty(region, conversationId, 'alerts', baseUrl),
    {
      method: 'PUT',
      headers: getMessagingHeaders(auth.skypeToken, auth.authToken, baseUrl),
      body: JSON.stringify({ alerts: muted ? 'false' : 'true' }),
    }
  );
  if (!response.ok) return response;
  return ok({ conversationId, muted });
}
