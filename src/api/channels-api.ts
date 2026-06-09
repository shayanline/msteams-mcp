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
  // needs the group GUID. The teams list exposes the thread id at team level and
  // the group GUID on each channel, so match on either to support being given a
  // group GUID (from teams_find_channel) or a team thread id.
  const teamsList = await getMyTeamsAndChannels();
  if (!teamsList.ok) {
    return teamsList;
  }
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
  const groupId = team.channels.find(c => c.teamId)?.teamId || teamId;

  const url = CHANNELS_API.createChannel(
    regionConfig.regionPartition,
    regionConfig.hasPartition,
    team.threadId,
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
