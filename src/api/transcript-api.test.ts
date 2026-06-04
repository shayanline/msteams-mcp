import { describe, it, expect, vi, beforeEach } from 'vitest';
import { ok } from '../types/result.js';

vi.mock('../utils/http.js', () => ({ httpRequest: vi.fn() }));
vi.mock('../utils/auth-guards.js', () => ({ requireSubstrateTokenAsync: vi.fn() }));

import { httpRequest } from '../utils/http.js';
import { requireSubstrateTokenAsync } from '../utils/auth-guards.js';
import { getTranscriptContent } from './transcript-api.js';

const mockHttp = vi.mocked(httpRequest);
const mockAuth = vi.mocked(requireSubstrateTokenAsync);
const httpOk = (data: unknown) => ok({ status: 200, headers: new Headers(), data } as never);
const fileItem = (transcriptJson: string | undefined) => ({
  ItemProperties: { Default: { TranscriptJson: transcriptJson, RecordingStartDateTime: 's', RecordingEndDateTime: 'e' } },
  Visualization: { Title: 'Standup' },
});

beforeEach(() => {
  vi.clearAllMocks();
  mockAuth.mockResolvedValue(ok('sub-token') as never);
});

describe('getTranscriptContent', () => {
  it('parses transcript entries and derives speakers', async () => {
    const tj = JSON.stringify({ entries: [
      { startOffset: '00:00:010000', endOffset: '00:05:010000', speakerDisplayName: 'Alice', text: 'Hello' },
      { startOffset: '00:06:000000', endOffset: '00:09:000000', speakerDisplayName: 'Bob', text: 'Hi' },
    ] });
    mockHttp.mockResolvedValueOnce(httpOk({ value: [fileItem(tj)] }));
    const res = await getTranscriptContent('19:meeting@thread.v2', '2026-06-04T10:00:00Z');
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.entryCount).toBe(2);
    expect(res.value.speakers).toEqual(['Alice', 'Bob']);
    expect(res.value.entries[0].startTime).toBe('00:00:01'); // trailing micro zeros trimmed
    expect(res.value.meetingTitle).toBe('Standup');
    expect(res.value.formattedText).toContain('Alice');
  });

  it('errors when no recording item is found', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ value: [] }));
    const res = await getTranscriptContent('19:m@thread.v2');
    expect(res.ok).toBe(false);
  });

  it('errors when the item has no TranscriptJson', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ value: [fileItem(undefined)] }));
    const res = await getTranscriptContent('19:m@thread.v2');
    expect(res.ok).toBe(false);
  });

  it('errors when the transcript JSON is invalid', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ value: [fileItem('not json')] }));
    const res = await getTranscriptContent('19:m@thread.v2');
    expect(res.ok).toBe(false);
  });

  it('errors when the transcript has no entries', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ value: [fileItem(JSON.stringify({ entries: [] }))] }));
    const res = await getTranscriptContent('19:m@thread.v2');
    expect(res.ok).toBe(false);
  });

  it('propagates auth failure', async () => {
    mockAuth.mockResolvedValueOnce({ ok: false, error: { code: 'AUTH_REQUIRED' } } as never);
    const res = await getTranscriptContent('19:m@thread.v2');
    expect(res.ok).toBe(false);
    expect(mockHttp).not.toHaveBeenCalled();
  });
});
