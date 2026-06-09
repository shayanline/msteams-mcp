/**
 * Channels API client for Teams channel management (middleTier).
 *
 * Handles calls to teams.microsoft.com/api/mt/part/{region}/beta/teams/{teamThreadId}/channels.
 * Creating a channel needs the team's ROOT thread id in the path (resolved from
 * the group GUID via the CSA teams list) and the group GUID in the body.
 */

import { httpRequest } from '../utils/http.js';
import { CHANNELS_API, getSkypeAuthHeaders } from '../utils/api-config.js';
import { type Result, ok, err } from '../types/result.js';
import { ErrorCode, createError } from '../types/errors.js';
import { requireSkypeSpacesAuth, getRegionConfig } from '../utils/auth-guards.js';
import { getMyTeamsAndChannels } from './csa-api.js';

/** Result of creating a channel. */
export interface CreateChannelResult {
  /** The new channel's conversation ID (19:...@thread.tacv2). */
  conversationId: string;
  displayName: string;
  description?: string;
  /** The host team's group ID (GUID). */
  teamId: string;
  membershipType: 'standard' | 'private';
}

/** Result of deleting a channel. */
export interface DeleteChannelResult {
  /** The deleted channel's conversation ID. */
  conversationId: string;
  /** The host team's group ID (GUID). */
  teamId: string;
}

/**
 * Resolves a team from a group GUID or team thread id to the identifiers the
 * mt/part channel endpoints need: the team root thread id (URL path) and the
 * group GUID (request body). The teams list exposes the thread id at team level
 * and the group GUID on each channel.
 */
async function resolveTeam(teamId: string): Promise<Result<{ threadId: string; groupId: string }>> {
  const teamsList = await getMyTeamsAndChannels();
  if (!teamsList.ok) return teamsList;
  const team = teamsList.value.teams.find(t =>
    t.threadId === teamId || t.teamId === teamId || t.channels.some(c => c.teamId === teamId)
  );
  if (!team) {
    return err(createError(
      ErrorCode.NOT_FOUND,
      `You are not a member of a team identified by "${teamId}", or it could not be found.`,
      { suggestions: ['Use teams_find_channel to get the correct teamId (group ID)'] }
    ));
  }
  return ok({ threadId: team.threadId, groupId: team.channels.find(c => c.teamId)?.teamId || teamId });
}

/**
 * Creates a channel in a team.
 *
 * @param teamId - The host team's group ID (GUID), as returned by teams_find_channel.
 * @param displayName - Channel name (max 50 chars).
 * @param description - Optional channel description.
 * @param membershipType - "standard" (default) or "private".
 */
export async function createChannel(
  teamId: string,
  displayName: string,
  description = '',
  membershipType: 'standard' | 'private' = 'standard'
): Promise<Result<CreateChannelResult>> {
  const authResult = requireSkypeSpacesAuth();
  if (!authResult.ok) {
    return authResult;
  }
  const { skypeToken, spacesToken } = authResult.value;

  const regionConfig = getRegionConfig();
  if (!regionConfig) {
    return err(createError(
      ErrorCode.AUTH_REQUIRED,
      'Could not determine region. Please run teams_login to authenticate.',
      { suggestions: ['Call teams_login to authenticate'] }
    ));
  }

  // The create-channel URL keys on the team's root thread id, while the body
  // needs the group GUID.
  const resolved = await resolveTeam(teamId);
  if (!resolved.ok) return resolved;
  const { threadId, groupId } = resolved.value;

  const url = CHANNELS_API.createChannel(
    regionConfig.regionPartition,
    regionConfig.hasPartition,
    threadId,
    regionConfig.teamsBaseUrl
  );

  const response = await httpRequest<Record<string, unknown>>(
    url,
    {
      method: 'POST',
      headers: getSkypeAuthHeaders(skypeToken, spacesToken, regionConfig.teamsBaseUrl),
      body: JSON.stringify({
        displayName,
        description,
        groupId,
        channelType: membershipType === 'private' ? 'Private' : 'Standard',
        chatModalityType: 'Conversational',
      }),
    }
  );

  if (!response.ok) {
    return response;
  }

  const value = (response.value.data.value ?? response.value.data) as Record<string, unknown>;
  const conversationId = value.objectId as string | undefined;
  if (!conversationId) {
    return err(createError(ErrorCode.API_ERROR, 'Channel was created but no conversation ID was returned.'));
  }

  return ok({
    conversationId,
    displayName: (value.displayName as string) || displayName,
    description: (value.description as string) || undefined,
    teamId: groupId,
    membershipType,
  });
}

/**
 * Deletes a channel from a team.
 *
 * @param teamId - The host team's group ID (GUID), as returned by teams_find_channel.
 * @param channelId - The channel's conversation ID (19:...@thread.tacv2).
 */
export async function deleteChannel(
  teamId: string,
  channelId: string
): Promise<Result<DeleteChannelResult>> {
  const authResult = requireSkypeSpacesAuth();
  if (!authResult.ok) {
    return authResult;
  }
  const { skypeToken, spacesToken } = authResult.value;

  const regionConfig = getRegionConfig();
  if (!regionConfig) {
    return err(createError(
      ErrorCode.AUTH_REQUIRED,
      'Could not determine region. Please run teams_login to authenticate.',
      { suggestions: ['Call teams_login to authenticate'] }
    ));
  }

  const resolved = await resolveTeam(teamId);
  if (!resolved.ok) return resolved;

  const url = CHANNELS_API.channel(
    regionConfig.regionPartition,
    regionConfig.hasPartition,
    resolved.value.threadId,
    channelId,
    regionConfig.teamsBaseUrl
  );

  const response = await httpRequest<unknown>(
    url,
    {
      method: 'DELETE',
      headers: getSkypeAuthHeaders(skypeToken, spacesToken, regionConfig.teamsBaseUrl),
    }
  );

  if (!response.ok) {
    return response;
  }

  return ok({ conversationId: channelId, teamId: resolved.value.groupId });
}
