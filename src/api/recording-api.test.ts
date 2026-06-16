import { describe, it, expect, vi, beforeEach } from 'vitest';
import { ok } from '../types/result.js';

vi.mock('../utils/http.js', () => ({ httpRequest: vi.fn() }));
vi.mock('../utils/auth-guards.js', () => ({ requireSubstrateTokenAsync: vi.fn() }));

import { httpRequest } from '../utils/http.js';
import { requireSubstrateTokenAsync } from '../utils/auth-guards.js';
import { getMeetingRecordings } from './recording-api.js';

const mockHttp = vi.mocked(httpRequest);
const mockAuth = vi.mocked(requireSubstrateTokenAsync);
const httpOk = (data: unknown) => ok({ status: 200, headers: new Headers(), data } as never);

const recordingItem = () => ({
  FileName: 'Standup-20260615-Meeting Recording.mp4',
  FileExtension: 'mp4',
  SharePointItem: {
    MediaDuration: 2571,
    FileUrl: 'https://contoso-my.sharepoint.com/personal/x/Documents/Recordings/Standup.mp4',
    DefaultEncodingUrl: 'https://contoso-my.sharepoint.com/personal/x/Documents/Recordings/Standup.mp4',
  },
  Visualization: { Title: 'Standup', Type: 'Video', AccessUrl: 'https://contoso-my.sharepoint.com/play' },
  ItemProperties: {
    Default: {
      MeetingThreadId: '19:meeting@thread.v2',
      MeetingCallId: 'call-1',
      DriveId: 'drive-1',
      DocumentLink: 'https://contoso-my.sharepoint.com/personal/x/Documents/Recordings/Standup.mp4',
      RecordingStartDateTime: '2026-06-15T08:32:47Z',
      RecordingEndDateTime: '2026-06-15T09:15:38Z',
    },
  },
});

beforeEach(() => {
  vi.clearAllMocks();
  mockAuth.mockResolvedValue(ok('sub-token') as never);
});

describe('getMeetingRecordings', () => {
  it('maps a recording item to playback/download URLs and metadata', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ value: [recordingItem()] }));
    const res = await getMeetingRecordings('19:meeting@thread.v2', '2026-06-15T08:00:00Z');
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.count).toBe(1);
    expect(res.value.meetingTitle).toBe('Standup');
    const rec = res.value.recordings[0];
    expect(rec.fileExtension).toBe('mp4');
    expect(rec.durationSeconds).toBe(2571);
    expect(rec.playbackUrl).toContain('Standup.mp4');
    expect(rec.downloadUrl).toContain('Standup.mp4');
    expect(rec.driveId).toBe('drive-1');
    expect(rec.callId).toBe('call-1');
    expect(rec.recordingStartTime).toBe('2026-06-15T08:32:47Z');
  });

  it('detects a recording by Visualization.Type=Video when no mp4 extension', async () => {
    const item = recordingItem();
    delete (item as Record<string, unknown>).FileExtension;
    mockHttp.mockResolvedValueOnce(httpOk({ value: [item] }));
    const res = await getMeetingRecordings('19:meeting@thread.v2');
    expect(res.ok).toBe(true);
  });

  it('ignores non-recording items (e.g. transcript-only artifacts)', async () => {
    const nonRec = { FileExtension: 'docx', Visualization: { Type: 'Document' } };
    mockHttp.mockResolvedValueOnce(httpOk({ value: [nonRec] }));
    const res = await getMeetingRecordings('19:meeting@thread.v2');
    expect(res.ok).toBe(false);
  });

  it('errors when no items are returned', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ value: [] }));
    const res = await getMeetingRecordings('19:m@thread.v2');
    expect(res.ok).toBe(false);
  });

  it('propagates auth failure without calling http', async () => {
    mockAuth.mockResolvedValueOnce({ ok: false, error: { code: 'AUTH_REQUIRED' } } as never);
    const res = await getMeetingRecordings('19:m@thread.v2');
    expect(res.ok).toBe(false);
    expect(mockHttp).not.toHaveBeenCalled();
  });

  it('propagates an http failure', async () => {
    mockHttp.mockResolvedValueOnce({ ok: false, error: { code: 'API_ERROR' } } as never);
    const res = await getMeetingRecordings('19:m@thread.v2');
    expect(res.ok).toBe(false);
  });

  it('handles an item missing ItemProperties/SharePointItem', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ value: [{ FileExtension: 'mp4' }] }));
    const res = await getMeetingRecordings('19:m@thread.v2');
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.recordings[0].durationSeconds).toBeUndefined();
    expect(res.value.recordings[0].playbackUrl).toBeUndefined();
  });
});
