/**
 * Unit tests for meeting/calendar tools (real handlers).
 */
import { describe, it, expect, vi, beforeEach } from 'vitest';
import { ok, err } from '../types/result.js';
import { createError, ErrorCode } from '../types/errors.js';
import './index.js';

vi.mock('../api/calendar-api.js', () => ({
  getCalendarView: vi.fn(),
  createMeeting: vi.fn(),
  getMeeting: vi.fn(),
  updateMeeting: vi.fn(),
  cancelMeeting: vi.fn(),
  respondToMeeting: vi.fn(),
  getSchedule: vi.fn(),
  findMeetingTimes: vi.fn(),
}));
vi.mock('../api/transcript-api.js', () => ({ getTranscriptContent: vi.fn() }));

import {
  getCalendarView,
  createMeeting,
  getMeeting,
  updateMeeting,
  cancelMeeting,
  respondToMeeting,
  getSchedule,
  findMeetingTimes,
} from '../api/calendar-api.js';
import { getTranscriptContent } from '../api/transcript-api.js';
import {
  getMeetingsTool,
  createMeetingTool,
  getMeetingTool,
  updateMeetingTool,
  cancelMeetingTool,
  respondToMeetingTool,
  getScheduleTool,
  findMeetingTimesTool,
  getTranscriptTool,
} from './meeting-tools.js';

const ctx = { server: {} } as never;
const anErr = err(createError(ErrorCode.API_ERROR, 'boom'));

beforeEach(() => vi.clearAllMocks());

describe('getMeetingsTool', () => {
  it('returns meetings on success', async () => {
    vi.mocked(getCalendarView).mockResolvedValue(ok({ count: 1, meetings: [{ id: 'e1' }] }) as never);
    const res = await getMeetingsTool.handler({ limit: 50 }, ctx);
    expect(res.success).toBe(true);
    if (res.success) expect(res.data).toEqual({ count: 1, meetings: [{ id: 'e1' }] });
  });
  it('propagates errors', async () => {
    vi.mocked(getCalendarView).mockResolvedValue(anErr as never);
    expect((await getMeetingsTool.handler({ limit: 50 }, ctx)).success).toBe(false);
  });
});

describe('createMeetingTool', () => {
  it('returns created meeting fields', async () => {
    vi.mocked(createMeeting).mockResolvedValue(
      ok({ id: 'e1', subject: 's', startTime: 'a', endTime: 'b', joinUrl: 'u' }) as never
    );
    const res = await createMeetingTool.handler(
      { subject: 's', startTime: 'a', endTime: 'b', isOnlineMeeting: true }, ctx
    );
    expect(res.success).toBe(true);
    if (res.success) expect(res.data.id).toBe('e1');
  });
  it('propagates errors', async () => {
    vi.mocked(createMeeting).mockResolvedValue(anErr as never);
    const res = await createMeetingTool.handler(
      { subject: 's', startTime: 'a', endTime: 'b', isOnlineMeeting: true }, ctx
    );
    expect(res.success).toBe(false);
  });
});

describe('getMeetingTool', () => {
  it('returns the meeting', async () => {
    vi.mocked(getMeeting).mockResolvedValue(ok({ id: 'e1' }) as never);
    const res = await getMeetingTool.handler({ eventId: 'e1' }, ctx);
    expect(res.success).toBe(true);
    if (res.success) expect(res.data).toEqual({ meeting: { id: 'e1' } });
  });
  it('propagates errors', async () => {
    vi.mocked(getMeeting).mockResolvedValue(anErr as never);
    expect((await getMeetingTool.handler({ eventId: 'e1' }, ctx)).success).toBe(false);
  });
});

describe('updateMeetingTool', () => {
  it('updates and returns the meeting', async () => {
    vi.mocked(updateMeeting).mockResolvedValue(ok({ id: 'e1', subject: 'new' }) as never);
    const res = await updateMeetingTool.handler({ eventId: 'e1', subject: 'new' }, ctx);
    expect(res.success).toBe(true);
    expect(vi.mocked(updateMeeting)).toHaveBeenCalledWith('e1', { subject: 'new' });
  });
  it('propagates errors', async () => {
    vi.mocked(updateMeeting).mockResolvedValue(anErr as never);
    expect((await updateMeetingTool.handler({ eventId: 'e1' }, ctx)).success).toBe(false);
  });
});

describe('cancelMeetingTool', () => {
  it('returns success', async () => {
    vi.mocked(cancelMeeting).mockResolvedValue(ok({}) as never);
    const res = await cancelMeetingTool.handler({ eventId: 'e1' }, ctx);
    expect(res.success).toBe(true);
    if (res.success) expect(res.data).toEqual({ success: true });
  });
  it('propagates errors', async () => {
    vi.mocked(cancelMeeting).mockResolvedValue(anErr as never);
    expect((await cancelMeetingTool.handler({ eventId: 'e1' }, ctx)).success).toBe(false);
  });
});

describe('respondToMeetingTool', () => {
  it('responds successfully', async () => {
    vi.mocked(respondToMeeting).mockResolvedValue(ok({}) as never);
    const res = await respondToMeetingTool.handler(
      { eventId: 'e1', response: 'accept', sendResponse: true }, ctx
    );
    expect(res.success).toBe(true);
  });
  it('propagates errors', async () => {
    vi.mocked(respondToMeeting).mockResolvedValue(anErr as never);
    const res = await respondToMeetingTool.handler(
      { eventId: 'e1', response: 'decline', sendResponse: true }, ctx
    );
    expect(res.success).toBe(false);
  });
});

describe('getScheduleTool', () => {
  it('returns schedules', async () => {
    vi.mocked(getSchedule).mockResolvedValue(ok({ schedules: [{ email: 'a@b.com' }] }) as never);
    const res = await getScheduleTool.handler(
      { schedules: ['a@b.com'], startTime: 'a', endTime: 'b', availabilityViewInterval: 30 }, ctx
    );
    expect(res.success).toBe(true);
    if (res.success) expect(res.data).toEqual({ schedules: [{ email: 'a@b.com' }] });
  });
  it('propagates errors', async () => {
    vi.mocked(getSchedule).mockResolvedValue(anErr as never);
    const res = await getScheduleTool.handler(
      { schedules: ['a@b.com'], startTime: 'a', endTime: 'b', availabilityViewInterval: 30 }, ctx
    );
    expect(res.success).toBe(false);
  });
});

describe('findMeetingTimesTool', () => {
  it('returns suggestions', async () => {
    vi.mocked(findMeetingTimes).mockResolvedValue(ok({ suggestions: [{}], emptyReason: undefined }) as never);
    const res = await findMeetingTimesTool.handler(
      { attendees: ['a@b.com'], durationMinutes: 30, maxCandidates: 5 }, ctx
    );
    expect(res.success).toBe(true);
    if (res.success) expect(res.data.suggestions).toEqual([{}]);
  });
  it('propagates errors', async () => {
    vi.mocked(findMeetingTimes).mockResolvedValue(anErr as never);
    const res = await findMeetingTimesTool.handler(
      { attendees: ['a@b.com'], durationMinutes: 30, maxCandidates: 5 }, ctx
    );
    expect(res.success).toBe(false);
  });
});

describe('getTranscriptTool', () => {
  it('returns transcript content', async () => {
    vi.mocked(getTranscriptContent).mockResolvedValue(
      ok({ meetingTitle: 't', speakers: ['a'], entryCount: 2, formattedText: 'text' }) as never
    );
    const res = await getTranscriptTool.handler({ threadId: '19:meet@thread.v2' }, ctx);
    expect(res.success).toBe(true);
    if (res.success) expect(res.data.transcript).toBe('text');
  });
  it('propagates errors', async () => {
    vi.mocked(getTranscriptContent).mockResolvedValue(anErr as never);
    expect((await getTranscriptTool.handler({ threadId: 'x' }, ctx)).success).toBe(false);
  });
});
