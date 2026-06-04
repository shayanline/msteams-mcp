import { describe, it, expect, vi, beforeEach } from 'vitest';
import { ok, err } from '../types/result.js';
import { ErrorCode, createError } from '../types/errors.js';

vi.mock('../utils/http.js', () => ({ httpRequest: vi.fn() }));
vi.mock('../utils/auth-guards.js', () => ({
  requireSkypeSpacesAuth: vi.fn(),
  requireGraphAuth: vi.fn(),
  getRegionConfig: vi.fn(),
}));

import { httpRequest } from '../utils/http.js';
import {
  requireSkypeSpacesAuth,
  requireGraphAuth,
  getRegionConfig,
} from '../utils/auth-guards.js';
import {
  toUtcIso,
  getCalendarView,
  createMeeting,
  getMeeting,
  updateMeeting,
  cancelMeeting,
  respondToMeeting,
  getSchedule,
  findMeetingTimes,
} from './calendar-api.js';

const mockHttp = vi.mocked(httpRequest);
const mockSpaces = vi.mocked(requireSkypeSpacesAuth);
const mockGraph = vi.mocked(requireGraphAuth);
const mockRegion = vi.mocked(getRegionConfig);

const httpOk = (data: unknown) => ok({ status: 200, headers: new Headers(), data } as never);
const httpErr = err(createError(ErrorCode.NETWORK_ERROR, 'boom'));
const authErr = err(createError(ErrorCode.AUTH_REQUIRED, 'no auth'));

beforeEach(() => {
  vi.clearAllMocks();
  mockSpaces.mockReturnValue(ok({ skypeToken: 'sk', spacesToken: 'sp' }) as never);
  mockGraph.mockReturnValue(ok('graph-token') as never);
  mockRegion.mockReturnValue({
    regionPartition: 'amer-02',
    hasPartition: true,
    teamsBaseUrl: 'https://teams.microsoft.com',
  } as never);
});

// ─────────────────────────────────────────────────────────────────────────────
// toUtcIso
// ─────────────────────────────────────────────────────────────────────────────

describe('toUtcIso', () => {
  it('returns empty string when no dateTime', () => {
    expect(toUtcIso(undefined)).toBe('');
    expect(toUtcIso({})).toBe('');
  });

  it('returns the raw value untouched when it already has Z or an offset', () => {
    expect(toUtcIso({ dateTime: '2026-06-05T10:00:00Z' })).toBe('2026-06-05T10:00:00Z');
    expect(toUtcIso({ dateTime: '2026-06-05T10:00:00+02:00' })).toBe('2026-06-05T10:00:00+02:00');
  });

  it('appends Z when timezone is UTC or absent', () => {
    expect(toUtcIso({ dateTime: '2026-06-05T10:00:00' })).toBe('2026-06-05T10:00:00Z');
    expect(toUtcIso({ dateTime: '2026-06-05T10:00:00', timeZone: 'UTC' })).toBe('2026-06-05T10:00:00Z');
  });

  it('leaves non-UTC values untouched', () => {
    expect(toUtcIso({ dateTime: '2026-06-05T10:00:00', timeZone: 'Pacific Standard Time' })).toBe('2026-06-05T10:00:00');
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// getCalendarView
// ─────────────────────────────────────────────────────────────────────────────

describe('getCalendarView', () => {
  it('propagates auth failure without calling http', async () => {
    mockSpaces.mockReturnValueOnce(authErr as never);
    const res = await getCalendarView();
    expect(res.ok).toBe(false);
    expect(mockHttp).not.toHaveBeenCalled();
  });

  it('errors when region config is null', async () => {
    mockRegion.mockReturnValueOnce(null);
    const res = await getCalendarView();
    expect(res.ok).toBe(false);
    if (res.ok) return;
    expect(res.error.code).toBe(ErrorCode.AUTH_REQUIRED);
    expect(mockHttp).not.toHaveBeenCalled();
  });

  it('parses meetings with all field variations and uses provided options', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({
      value: [
        {
          objectId: 'm1',
          subject: 'Sync',
          startTime: '2026-06-05T10:00:00Z',
          endTime: '2026-06-05T10:30:00Z',
          organizerName: 'Alice',
          organizerAddress: 'alice@x.com',
          location: 'Room 1',
          isOnlineMeeting: true,
          skypeTeamsMeetingUrl: 'https://join',
          skypeTeamsData: JSON.stringify({ cid: '19:thread@thread.v2' }),
          schedulingServiceUpdateUrl: 'https://upd',
          myResponseType: 'Accepted',
          showAs: 'Busy',
          isOrganizer: true,
          eventType: 'Occurrence',
        },
      ],
    }));
    const res = await getCalendarView({
      startDate: '2026-06-01T00:00:00Z',
      endDate: '2026-06-08T00:00:00Z',
      limit: 10,
    });
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.count).toBe(1);
    const m = res.value.meetings[0];
    expect(m).toMatchObject({
      id: 'm1',
      subject: 'Sync',
      threadId: '19:thread@thread.v2',
      myResponse: 'Accepted',
      showAs: 'Busy',
      isOrganizer: true,
      isOnlineMeeting: true,
      eventType: 'Occurrence',
    });
    // url contains provided startDate / limit / region partition
    const url = mockHttp.mock.calls[0][0] as string;
    expect(url).toContain('startDate=2026-06-01');
    expect(url).toContain('%24top=10');
  });

  it('applies fallbacks and maps Organizer/Tentative/Oof/declined variants and bad JSON', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({
      value: [
        {
          objectId: 'm2',
          // no subject → '(No subject)', no organizer fields → Unknown/'',
          skypeTeamsData: 'not-json', // triggers catch
          myResponseType: 'Organizer',
          showAs: 'Free',
        },
        {
          objectId: 'm3',
          myResponseType: 'TentativelyAccepted',
          showAs: 'Tentative',
        },
        {
          objectId: 'm4',
          myResponseType: 'Declined',
          showAs: 'Oof',
        },
        {
          objectId: 'm5',
          myResponseType: 'Unknown',
          showAs: 'OutOfOffice',
        },
        {
          objectId: 'm6',
          myResponseType: 'Tentative',
          showAs: 'somethingElse',
        },
      ],
    }));
    const res = await getCalendarView();
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    const [a, b, c, d, e] = res.value.meetings;
    expect(a.subject).toBe('(No subject)');
    expect(a.organizer).toEqual({ name: 'Unknown', email: '' });
    expect(a.threadId).toBeUndefined();
    expect(a.myResponse).toBe('Accepted');
    expect(a.showAs).toBe('Free');
    expect(a.eventType).toBe('Single');
    expect(b).toMatchObject({ myResponse: 'Tentative', showAs: 'Tentative' });
    expect(c).toMatchObject({ myResponse: 'Declined', showAs: 'OutOfOffice' });
    expect(d).toMatchObject({ myResponse: 'None', showAs: 'OutOfOffice' });
    expect(e).toMatchObject({ myResponse: 'Tentative', showAs: 'Unknown' });
  });

  it('returns empty when the response has no value array', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({}));
    const res = await getCalendarView();
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value).toEqual({ count: 0, meetings: [] });
  });

  it('returns empty when value is not an array', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ value: 'nope' }));
    const res = await getCalendarView();
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.meetings).toEqual([]);
  });

  it('propagates an http failure', async () => {
    mockHttp.mockResolvedValueOnce(httpErr as never);
    const res = await getCalendarView();
    expect(res.ok).toBe(false);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// createMeeting
// ─────────────────────────────────────────────────────────────────────────────

describe('createMeeting', () => {
  it('propagates auth failure', async () => {
    mockGraph.mockReturnValueOnce(authErr as never);
    const res = await createMeeting({ subject: 's', startTime: 'a', endTime: 'b' });
    expect(res.ok).toBe(false);
    expect(mockHttp).not.toHaveBeenCalled();
  });

  it('creates an online meeting by default and builds full body', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({
      id: 'evt1',
      subject: 'Planning',
      start: { dateTime: '2026-06-05T10:00:00', timeZone: 'UTC' },
      end: { dateTime: '2026-06-05T11:00:00', timeZone: 'UTC' },
      onlineMeeting: { joinUrl: 'https://join.me' },
    }));
    const res = await createMeeting({
      subject: 'Planning',
      startTime: '2026-06-05T10:00:00Z',
      endTime: '2026-06-05T11:00:00Z',
      body: 'agenda',
      location: 'Room',
      attendees: [{ email: 'a@x.com', name: 'A' }, { email: 'b@x.com' }],
    });
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value).toEqual({
      id: 'evt1',
      subject: 'Planning',
      startTime: '2026-06-05T10:00:00Z',
      endTime: '2026-06-05T11:00:00Z',
      joinUrl: 'https://join.me',
    });
    const sent = JSON.parse((mockHttp.mock.calls[0][1] as { body: string }).body);
    expect(sent.isOnlineMeeting).toBe(true);
    expect(sent.onlineMeetingProvider).toBe('teamsForBusiness');
    expect(sent.attendees[0]).toMatchObject({ emailAddress: { address: 'a@x.com', name: 'A' } });
    expect(sent.attendees[1]).toMatchObject({ emailAddress: { address: 'b@x.com', name: 'b@x.com' } });
    expect(sent.body).toEqual({ contentType: 'text', content: 'agenda' });
    expect(sent.location).toEqual({ displayName: 'Room' });
  });

  it('respects isOnlineMeeting false and falls back to option values', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({})); // no id/subject/start/end
    const res = await createMeeting({
      subject: 'Offline',
      startTime: '2026-06-05T10:00:00Z',
      endTime: '2026-06-05T11:00:00Z',
      isOnlineMeeting: false,
    });
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.subject).toBe('Offline');
    expect(res.value.startTime).toBe('2026-06-05T10:00:00Z');
    expect(res.value.endTime).toBe('2026-06-05T11:00:00Z');
    expect(res.value.joinUrl).toBeUndefined();
    const sent = JSON.parse((mockHttp.mock.calls[0][1] as { body: string }).body);
    expect(sent.isOnlineMeeting).toBe(false);
    expect(sent.onlineMeetingProvider).toBeUndefined();
  });

  it('propagates an http failure', async () => {
    mockHttp.mockResolvedValueOnce(httpErr as never);
    const res = await createMeeting({ subject: 's', startTime: 'a', endTime: 'b' });
    expect(res.ok).toBe(false);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// getMeeting / updateMeeting (parseGraphEvent)
// ─────────────────────────────────────────────────────────────────────────────

describe('getMeeting', () => {
  it('propagates auth failure', async () => {
    mockGraph.mockReturnValueOnce(authErr as never);
    const res = await getMeeting('e1');
    expect(res.ok).toBe(false);
    expect(mockHttp).not.toHaveBeenCalled();
  });

  it('parses a full graph event', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({
      id: 'e1',
      subject: 'Review',
      start: { dateTime: '2026-06-05T10:00:00', timeZone: 'UTC' },
      end: { dateTime: '2026-06-05T11:00:00', timeZone: 'UTC' },
      organizer: { emailAddress: { name: 'Bob', address: 'bob@x.com' } },
      location: { displayName: 'HQ' },
      onlineMeeting: { joinUrl: 'https://j' },
      isOnlineMeeting: true,
      isOrganizer: true,
      responseStatus: { response: 'accepted' },
      showAs: 'busy',
      type: 'occurrence',
    }));
    const res = await getMeeting('e1');
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value).toMatchObject({
      id: 'e1',
      subject: 'Review',
      startTime: '2026-06-05T10:00:00Z',
      organizer: { name: 'Bob', email: 'bob@x.com' },
      location: 'HQ',
      joinUrl: 'https://j',
      myResponse: 'Accepted',
      showAs: 'Busy',
      eventType: 'occurrence',
    });
  });

  it('applies fallbacks and maps organizer/tentative/declined/workingElsewhere/free variants', async () => {
    mockHttp
      .mockResolvedValueOnce(httpOk({
        id: 'e2',
        responseStatus: { response: 'organizer' },
        showAs: 'free',
      }))
      .mockResolvedValueOnce(httpOk({
        id: 'e3',
        responseStatus: { response: 'tentativelyAccepted' },
        showAs: 'tentative',
      }))
      .mockResolvedValueOnce(httpOk({
        id: 'e4',
        responseStatus: { response: 'declined' },
        showAs: 'oof',
      }))
      .mockResolvedValueOnce(httpOk({
        id: 'e5',
        responseStatus: { response: 'notResponded' },
        showAs: 'workingElsewhere',
      }))
      .mockResolvedValueOnce(httpOk({
        id: 'e6',
        showAs: 'mystery',
      }));

    const r1 = await getMeeting('e2');
    const r2 = await getMeeting('e3');
    const r3 = await getMeeting('e4');
    const r4 = await getMeeting('e5');
    const r5 = await getMeeting('e6');
    if (!r1.ok || !r2.ok || !r3.ok || !r4.ok || !r5.ok) throw new Error('expected ok');
    expect(r1.value).toMatchObject({ subject: '(No subject)', myResponse: 'Accepted', showAs: 'Free', eventType: 'singleInstance' });
    expect(r1.value.organizer).toEqual({ name: 'Unknown', email: '' });
    expect(r2.value).toMatchObject({ myResponse: 'Tentative', showAs: 'Tentative' });
    expect(r3.value).toMatchObject({ myResponse: 'Declined', showAs: 'OutOfOffice' });
    expect(r4.value).toMatchObject({ myResponse: 'None', showAs: 'OutOfOffice' });
    expect(r5.value).toMatchObject({ myResponse: 'None', showAs: 'Unknown' });
  });

  it('propagates an http failure', async () => {
    mockHttp.mockResolvedValueOnce(httpErr as never);
    const res = await getMeeting('e1');
    expect(res.ok).toBe(false);
  });
});

describe('updateMeeting', () => {
  it('propagates auth failure', async () => {
    mockGraph.mockReturnValueOnce(authErr as never);
    const res = await updateMeeting('e1', { subject: 'x' });
    expect(res.ok).toBe(false);
    expect(mockHttp).not.toHaveBeenCalled();
  });

  it('PATCHes with a partial body and returns the parsed event', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ id: 'e1', subject: 'Updated' }));
    const res = await updateMeeting('e1', { subject: 'Updated' });
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.subject).toBe('Updated');
    expect(mockHttp.mock.calls[0][1]).toMatchObject({ method: 'PATCH' });
    const sent = JSON.parse((mockHttp.mock.calls[0][1] as { body: string }).body);
    expect(sent).toEqual({ subject: 'Updated' });
  });

  it('propagates an http failure', async () => {
    mockHttp.mockResolvedValueOnce(httpErr as never);
    const res = await updateMeeting('e1', { subject: 'x' });
    expect(res.ok).toBe(false);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// cancelMeeting
// ─────────────────────────────────────────────────────────────────────────────

describe('cancelMeeting', () => {
  it('propagates auth failure', async () => {
    mockGraph.mockReturnValueOnce(authErr as never);
    const res = await cancelMeeting('e1');
    expect(res.ok).toBe(false);
    expect(mockHttp).not.toHaveBeenCalled();
  });

  it('deletes and returns ok(undefined)', async () => {
    mockHttp.mockResolvedValueOnce(httpOk(undefined));
    const res = await cancelMeeting('e1');
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value).toBeUndefined();
    expect(mockHttp.mock.calls[0][1]).toMatchObject({ method: 'DELETE' });
  });

  it('propagates an http failure', async () => {
    mockHttp.mockResolvedValueOnce(httpErr as never);
    const res = await cancelMeeting('e1');
    expect(res.ok).toBe(false);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// respondToMeeting
// ─────────────────────────────────────────────────────────────────────────────

describe('respondToMeeting', () => {
  it('propagates auth failure', async () => {
    mockGraph.mockReturnValueOnce(authErr as never);
    const res = await respondToMeeting('e1', { response: 'accept' });
    expect(res.ok).toBe(false);
    expect(mockHttp).not.toHaveBeenCalled();
  });

  it('sends comment and proposedNewTime for a decline', async () => {
    mockHttp.mockResolvedValueOnce(httpOk(undefined));
    const res = await respondToMeeting('e1', {
      response: 'decline',
      comment: 'cannot make it',
      proposedNewTime: { start: '2026-06-06T10:00:00Z', end: '2026-06-06T11:00:00Z' },
    });
    expect(res.ok).toBe(true);
    const sent = JSON.parse((mockHttp.mock.calls[0][1] as { body: string }).body);
    expect(sent.sendResponse).toBe(true);
    expect(sent.comment).toBe('cannot make it');
    expect(sent.proposedNewTime.start).toEqual({ dateTime: '2026-06-06T10:00:00Z', timeZone: 'UTC' });
  });

  it('omits comment when sendResponse is false and omits proposedNewTime on accept', async () => {
    mockHttp.mockResolvedValueOnce(httpOk(undefined));
    const res = await respondToMeeting('e1', {
      response: 'accept',
      comment: 'thanks',
      sendResponse: false,
      proposedNewTime: { start: 'a', end: 'b' },
    });
    expect(res.ok).toBe(true);
    const sent = JSON.parse((mockHttp.mock.calls[0][1] as { body: string }).body);
    expect(sent.sendResponse).toBe(false);
    expect(sent.comment).toBeUndefined();
    expect(sent.proposedNewTime).toBeUndefined();
  });

  it('propagates an http failure', async () => {
    mockHttp.mockResolvedValueOnce(httpErr as never);
    const res = await respondToMeeting('e1', { response: 'tentativelyAccept' });
    expect(res.ok).toBe(false);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// getSchedule
// ─────────────────────────────────────────────────────────────────────────────

describe('getSchedule', () => {
  it('propagates auth failure', async () => {
    mockGraph.mockReturnValueOnce(authErr as never);
    const res = await getSchedule({ schedules: ['a@x.com'], startTime: 's', endTime: 'e' });
    expect(res.ok).toBe(false);
    expect(mockHttp).not.toHaveBeenCalled();
  });

  it('parses schedules with items and uses provided interval', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({
      value: [
        {
          scheduleId: 'a@x.com',
          availabilityView: '0022',
          scheduleItems: [
            { start: { dateTime: '2026-06-05T10:00:00', timeZone: 'UTC' }, end: { dateTime: '2026-06-05T10:30:00', timeZone: 'UTC' }, status: 'busy', subject: 'Call', isPrivate: false },
          ],
        },
      ],
    }));
    const res = await getSchedule({
      schedules: ['a@x.com'],
      startTime: '2026-06-05T00:00:00Z',
      endTime: '2026-06-06T00:00:00Z',
      availabilityViewInterval: 60,
    });
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.schedules[0].scheduleId).toBe('a@x.com');
    expect(res.value.schedules[0].scheduleItems[0]).toMatchObject({ start: '2026-06-05T10:00:00Z', status: 'busy', subject: 'Call' });
    const sent = JSON.parse((mockHttp.mock.calls[0][1] as { body: string }).body);
    expect(sent.availabilityViewInterval).toBe(60);
  });

  it('applies fallbacks for missing fields and default interval', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({
      value: [
        { scheduleId: 'b@x.com' }, // no availabilityView, no scheduleItems
        { scheduleId: 'c@x.com', scheduleItems: [{ start: {}, end: {} }] }, // missing status
      ],
    }));
    const res = await getSchedule({ schedules: ['b@x.com', 'c@x.com'], startTime: 's', endTime: 'e' });
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.schedules[0].availabilityView).toBe('');
    expect(res.value.schedules[0].scheduleItems).toEqual([]);
    expect(res.value.schedules[1].scheduleItems[0].status).toBe('unknown');
    const sent = JSON.parse((mockHttp.mock.calls[0][1] as { body: string }).body);
    expect(sent.availabilityViewInterval).toBe(30);
  });

  it('returns empty schedules when value is missing', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({}));
    const res = await getSchedule({ schedules: ['a@x.com'], startTime: 's', endTime: 'e' });
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.schedules).toEqual([]);
  });

  it('propagates an http failure', async () => {
    mockHttp.mockResolvedValueOnce(httpErr as never);
    const res = await getSchedule({ schedules: ['a@x.com'], startTime: 's', endTime: 'e' });
    expect(res.ok).toBe(false);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// findMeetingTimes
// ─────────────────────────────────────────────────────────────────────────────

describe('findMeetingTimes', () => {
  it('propagates auth failure', async () => {
    mockGraph.mockReturnValueOnce(authErr as never);
    const res = await findMeetingTimes({ attendees: ['a@x.com'] });
    expect(res.ok).toBe(false);
    expect(mockHttp).not.toHaveBeenCalled();
  });

  it('parses suggestions and uses provided options', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({
      meetingTimeSuggestions: [
        {
          meetingTimeSlot: { start: { dateTime: '2026-06-05T10:00:00' }, end: { dateTime: '2026-06-05T10:30:00' } },
          confidence: 100,
          organizerAvailability: 'free',
          attendeeAvailability: [
            { attendee: { emailAddress: { address: 'a@x.com' } }, availability: 'free' },
          ],
        },
      ],
      emptySuggestionsReason: '',
    }));
    const res = await findMeetingTimes({
      attendees: ['a@x.com'],
      durationMinutes: 45,
      start: '2026-06-05T00:00:00Z',
      end: '2026-06-06T00:00:00Z',
      maxCandidates: 3,
    });
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.suggestions[0]).toMatchObject({
      start: '2026-06-05T10:00:00',
      end: '2026-06-05T10:30:00',
      confidence: 100,
      organizerAvailability: 'free',
    });
    expect(res.value.suggestions[0].attendeeAvailability).toEqual([{ email: 'a@x.com', availability: 'free' }]);
    expect(res.value.emptyReason).toBeUndefined();
    const sent = JSON.parse((mockHttp.mock.calls[0][1] as { body: string }).body);
    expect(sent.meetingDuration).toBe('PT45M');
    expect(sent.maxCandidates).toBe(3);
  });

  it('applies fallbacks for missing fields and default options', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({
      meetingTimeSuggestions: [
        { attendeeAvailability: [{}] }, // no slot, no confidence, attendee missing emailAddress
      ],
      emptySuggestionsReason: 'AttendeesUnavailable',
    }));
    const res = await findMeetingTimes({ attendees: ['a@x.com'] });
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.suggestions[0]).toMatchObject({ start: '', end: '' });
    expect(res.value.suggestions[0].attendeeAvailability).toEqual([{ email: '', availability: 'unknown' }]);
    expect(res.value.emptyReason).toBe('AttendeesUnavailable');
    const sent = JSON.parse((mockHttp.mock.calls[0][1] as { body: string }).body);
    expect(sent.meetingDuration).toBe('PT30M');
    expect(sent.maxCandidates).toBe(5);
  });

  it('returns empty suggestions when none are present', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({}));
    const res = await findMeetingTimes({ attendees: ['a@x.com'] });
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.suggestions).toEqual([]);
  });

  it('propagates an http failure', async () => {
    mockHttp.mockResolvedValueOnce(httpErr as never);
    const res = await findMeetingTimes({ attendees: ['a@x.com'] });
    expect(res.ok).toBe(false);
  });
});
