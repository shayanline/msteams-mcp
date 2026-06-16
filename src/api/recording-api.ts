/**
 * Recording API client for meeting recording operations.
 *
 * Uses the same Substrate WorkingSetFiles API as transcripts: each recorded
 * meeting surfaces an .mp4 artifact keyed by MeetingThreadId. We return its
 * playback/download URLs and metadata (no binary content — recordings are large
 * video files served from SharePoint/OneDrive).
 *
 * Flow: threadId → Substrate WorkingSetFiles (filter by MeetingThreadId) → mp4 items
 */

import { httpRequest } from '../utils/http.js';
import { type Result, ok, err } from '../types/result.js';
import { ErrorCode, createError } from '../types/errors.js';
import { requireSubstrateTokenAsync } from '../utils/auth-guards.js';

// ─────────────────────────────────────────────────────────────────────────────
// Types
// ─────────────────────────────────────────────────────────────────────────────

/** A single meeting recording file. */
export interface RecordingFile {
  /** Recording title (usually the meeting subject). */
  title?: string;
  /** File name, e.g. "RPA Migration Plan-...-Meeting Recording.mp4". */
  fileName?: string;
  /** File extension, e.g. "mp4". */
  fileExtension?: string;
  /** When the recording started. */
  recordingStartTime?: string;
  /** When the recording ended. */
  recordingEndTime?: string;
  /** Recording length in seconds, if known. */
  durationSeconds?: number;
  /** URL that opens the recording in the SharePoint/Stream web player. */
  playbackUrl?: string;
  /** Direct URL to the .mp4 file in SharePoint/OneDrive. */
  downloadUrl?: string;
  /** SharePoint drive ID hosting the recording. */
  driveId?: string;
  /** Underlying meeting call ID. */
  callId?: string;
}

/** Result of fetching meeting recordings. */
export interface RecordingResult {
  /** Meeting thread ID used for the lookup. */
  threadId: string;
  /** Meeting title from the recording metadata. */
  meetingTitle?: string;
  /** Recordings found for the meeting (most recent first). */
  recordings: RecordingFile[];
  /** Number of recordings found. */
  count: number;
}

// ─────────────────────────────────────────────────────────────────────────────
// Constants
// ─────────────────────────────────────────────────────────────────────────────

/** Substrate WorkingSetFiles endpoint for finding meeting recordings. */
const WORKING_SET_FILES_URL = 'https://substrate.office.com/api/beta/me/WorkingSetFiles/';

/** Fields to select from WorkingSetFiles for recordings. */
const WORKING_SET_SELECT = [
  'FileName',
  'FileExtension',
  'SharePointItem',
  'Visualization',
  'ItemProperties/Default/MeetingThreadId',
  'ItemProperties/Default/MeetingCallId',
  'ItemProperties/Default/DriveId',
  'ItemProperties/Default/DocumentLink',
  'ItemProperties/Default/RecordingStartDateTime',
  'ItemProperties/Default/RecordingEndDateTime',
].join(',');

// ─────────────────────────────────────────────────────────────────────────────
// API Functions
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Gets the recording(s) for a meeting by its thread ID.
 *
 * @param threadId - The meeting thread ID (e.g., "19:meeting_xxx@thread.v2")
 * @param meetingDate - ISO date string of the meeting (used to narrow the search window)
 * @returns Recording metadata with playback and download URLs
 */
export async function getMeetingRecordings(
  threadId: string,
  meetingDate?: string
): Promise<Result<RecordingResult>> {
  const authResult = await requireSubstrateTokenAsync();
  if (!authResult.ok) return authResult;
  const token = authResult.value;

  // Build date filter: search ±1 day around the meeting date when provided.
  let dateFilter = '';
  if (meetingDate) {
    const date = new Date(meetingDate);
    const dayBefore = new Date(date);
    dayBefore.setDate(dayBefore.getDate() - 1);
    const dayAfter = new Date(date);
    dayAfter.setDate(dayAfter.getDate() + 1);
    dateFilter = ` AND FileCreatedTime gt ${dayBefore.toISOString()} AND FileCreatedTime lt ${dayAfter.toISOString()}`;
  }

  const filter = `ItemProperties/Default/MeetingThreadId eq '${threadId}'${dateFilter}`;
  const url = `${WORKING_SET_FILES_URL}?$filter=${encodeURIComponent(filter)}&$orderby=${encodeURIComponent('FileCreatedTime desc')}&$select=${encodeURIComponent(WORKING_SET_SELECT)}`;

  const response = await httpRequest<Record<string, unknown>>(
    url,
    {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Accept': 'application/json',
        'Content-Type': 'application/json',
        'Prefer': 'substrate.flexibleschema,outlook.data-source="Substrate",exchange.behavior="SubstrateFiles"',
      },
    }
  );

  if (!response.ok) {
    return response;
  }

  const items = (response.value.data.value as Array<Record<string, unknown>> | undefined) ?? [];

  const recordings: RecordingFile[] = items
    .filter(isRecordingItem)
    .map(toRecordingFile);

  if (recordings.length === 0) {
    return err(createError(
      ErrorCode.NOT_FOUND,
      'No recording found for this meeting. The meeting may not have been recorded, or the recording has not finished processing.',
      { suggestions: [
        'Check that recording was enabled during the meeting',
        'Wait a few minutes if the meeting just ended',
        'Use teams_get_transcript if you only need the spoken content',
      ] }
    ));
  }

  return ok({
    threadId,
    meetingTitle: recordings[0]?.title,
    recordings,
    count: recordings.length,
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────────────────────────────────────────

/** True if a WorkingSetFiles item is a meeting recording (video / mp4). */
function isRecordingItem(item: Record<string, unknown>): boolean {
  const ext = (item.FileExtension as string | undefined)?.toLowerCase();
  const vizType = ((item.Visualization as Record<string, unknown> | undefined)?.Type as string | undefined)?.toLowerCase();
  return ext === 'mp4' || vizType === 'video';
}

/** Maps a raw WorkingSetFiles recording item to a RecordingFile. */
function toRecordingFile(item: Record<string, unknown>): RecordingFile {
  const props = (item.ItemProperties as Record<string, unknown> | undefined)?.Default as Record<string, unknown> | undefined;
  const viz = item.Visualization as Record<string, unknown> | undefined;
  const sp = item.SharePointItem as Record<string, unknown> | undefined;

  const duration = sp?.MediaDuration;

  return {
    title: viz?.Title as string | undefined,
    fileName: item.FileName as string | undefined,
    fileExtension: item.FileExtension as string | undefined,
    recordingStartTime: props?.RecordingStartDateTime as string | undefined,
    recordingEndTime: props?.RecordingEndDateTime as string | undefined,
    durationSeconds: typeof duration === 'number' ? duration : undefined,
    playbackUrl: (props?.DocumentLink as string | undefined) ?? (viz?.AccessUrl as string | undefined),
    downloadUrl: (sp?.FileUrl as string | undefined) ?? (sp?.DefaultEncodingUrl as string | undefined),
    driveId: props?.DriveId as string | undefined,
    callId: props?.MeetingCallId as string | undefined,
  };
}
