import { describe, it, expect, vi, beforeEach } from 'vitest';
import { ok, err } from '../types/result.js';
import { ErrorCode, createError } from '../types/errors.js';

vi.mock('../utils/http.js', () => ({ httpRequest: vi.fn() }));
vi.mock('../utils/auth-guards.js', () => ({ requireGraphAuth: vi.fn() }));
vi.mock('./chatsvc-messaging.js', () => ({ sendMessage: vi.fn(), getChannelFilesInfo: vi.fn() }));
vi.mock('node:fs/promises', () => ({ readFile: vi.fn(), writeFile: vi.fn() }));

import { httpRequest } from '../utils/http.js';
import { requireGraphAuth } from '../utils/auth-guards.js';
import { sendMessage, getChannelFilesInfo } from './chatsvc-messaging.js';
import { readFile, writeFile } from 'node:fs/promises';
import { listDriveFiles, uploadFile, downloadFile, createShareLink, sendFileToChat, sendFilesToChat } from './files-graph-api.js';

const mockHttp = vi.mocked(httpRequest);
const mockAuth = vi.mocked(requireGraphAuth);
const mockSend = vi.mocked(sendMessage);
const mockChannelInfo = vi.mocked(getChannelFilesInfo);
const mockRead = vi.mocked(readFile);
const mockWrite = vi.mocked(writeFile);

/** A drive item GET response shaped for getShareFileInfo (sharepointIds + webUrl). */
const shareInfoResponse = (webUrl: string) => httpOk({
  id: 'item1', name: 'report.pdf', webUrl,
  sharepointIds: { listItemUniqueId: 'LIU-GUID', siteId: 'SITE-GUID' },
});
const httpOk = (data: unknown) => ok({ status: 200, headers: new Headers(), data } as never);
const httpErr = (message: string) => err(createError(ErrorCode.UNKNOWN, message)) as never;

beforeEach(() => {
  vi.clearAllMocks();
  mockAuth.mockReturnValue(ok('graph-token') as never);
});

describe('listDriveFiles', () => {
  it('lists the root and maps folder flag', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ value: [
      { id: 'f1', name: 'Docs', folder: {} },
      { id: 'f2', name: 'a.txt', size: 10 },
    ] }));
    const res = await listDriveFiles();
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.items).toEqual([
      { id: 'f1', name: 'Docs', webUrl: undefined, size: undefined, isFolder: true },
      { id: 'f2', name: 'a.txt', webUrl: undefined, size: 10, isFolder: false },
    ]);
    expect(mockHttp.mock.calls[0][0]).toContain('/me/drive/root/children');
  });

  it('lists a sub-folder path', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ value: [] }));
    await listDriveFiles('Documents/Reports');
    expect(mockHttp.mock.calls[0][0]).toContain('/me/drive/root:/Documents/Reports:/children');
  });
});

describe('uploadFile', () => {
  it('PUTs file content to the given folder', async () => {
    mockRead.mockResolvedValueOnce(Buffer.from('hello'));
    mockHttp.mockResolvedValueOnce(httpOk({ id: '01X', name: 'note.txt', webUrl: 'http://x', size: 5 }));
    const res = await uploadFile('/tmp/note.txt', 'MyFolder');
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value).toMatchObject({ id: '01X', name: 'note.txt' });
    const call = mockHttp.mock.calls[0];
    expect(call[0]).toContain('/me/drive/root:/MyFolder/note.txt:/content');
    expect(call[1]).toMatchObject({ method: 'PUT' });
  });

  it('returns an input error when the file cannot be read', async () => {
    mockRead.mockRejectedValueOnce(new Error('ENOENT'));
    const res = await uploadFile('/tmp/missing.txt');
    expect(res.ok).toBe(false);
    expect(mockHttp).not.toHaveBeenCalled();
  });

  it('retries under a de-duplicated name when the target is locked (HTTP 423)', async () => {
    mockRead.mockResolvedValueOnce(Buffer.from('hello'));
    mockHttp
      .mockResolvedValueOnce(httpErr('HTTP 423: {"error":{"code":"resourceLocked"}}')) // original locked
      .mockResolvedValueOnce(httpOk({ id: '01Y', name: 'note (2).txt' }));              // retry succeeds
    const res = await uploadFile('/tmp/note.txt', 'MyFolder');
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value).toMatchObject({ id: '01Y', name: 'note (2).txt' });
    expect(mockHttp).toHaveBeenCalledTimes(2);
    expect(mockHttp.mock.calls[0][0]).toContain('/MyFolder/note.txt:/content');
    expect(mockHttp.mock.calls[1][0]).toContain('/MyFolder/note%20(2).txt:/content');
  });

  it('returns a retryable, well-signposted error when every name variant is locked', async () => {
    mockRead.mockResolvedValueOnce(Buffer.from('hello'));
    mockHttp.mockResolvedValue(httpErr('HTTP 423: resourceLocked'));
    const res = await uploadFile('/tmp/note.txt', 'MyFolder');
    expect(res.ok).toBe(false);
    if (res.ok) return;
    expect(mockHttp).toHaveBeenCalledTimes(4); // original + 3 fallbacks
    expect(res.error.retryable).toBe(true);
    expect(res.error.suggestions.join(' ').toLowerCase()).toContain('lock');
  });

  it('surfaces a non-lock upload failure immediately without renaming', async () => {
    mockRead.mockResolvedValueOnce(Buffer.from('hello'));
    mockHttp.mockResolvedValueOnce(httpErr('HTTP 403: Forbidden'));
    const res = await uploadFile('/tmp/note.txt', 'MyFolder');
    expect(res.ok).toBe(false);
    expect(mockHttp).toHaveBeenCalledTimes(1);
  });

  it('does not treat a non-423 error that merely mentions a lock as retryable', async () => {
    mockRead.mockResolvedValueOnce(Buffer.from('hello'));
    mockHttp.mockResolvedValueOnce(httpErr('HTTP 403: your account is locked'));
    const res = await uploadFile('/tmp/note.txt', 'MyFolder');
    expect(res.ok).toBe(false);
    expect(mockHttp).toHaveBeenCalledTimes(1); // surfaced immediately, no rename retries
  });
});

describe('downloadFile', () => {
  it('writes fetched bytes to the output path', async () => {
    const fetchMock = vi.fn().mockResolvedValue({ ok: true, arrayBuffer: async () => new TextEncoder().encode('data').buffer });
    vi.stubGlobal('fetch', fetchMock);
    const res = await downloadFile('01X', '/tmp/out.txt');
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.outputPath).toBe('/tmp/out.txt');
    expect(mockWrite).toHaveBeenCalledWith('/tmp/out.txt', expect.anything());
    vi.unstubAllGlobals();
  });

  it('errors on a non-ok download response', async () => {
    const fetchMock = vi.fn().mockResolvedValue({ ok: false, status: 404, text: async () => 'nope' });
    vi.stubGlobal('fetch', fetchMock);
    const res = await downloadFile('01X', '/tmp/out.txt');
    expect(res.ok).toBe(false);
    expect(mockWrite).not.toHaveBeenCalled();
    vi.unstubAllGlobals();
  });
});

describe('createShareLink', () => {
  it('returns the share webUrl', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({ link: { webUrl: 'https://share/x' } }));
    const res = await createShareLink('01X', 'edit');
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.webUrl).toBe('https://share/x');
    expect(JSON.parse((mockHttp.mock.calls[0][1] as { body: string }).body)).toEqual({ type: 'edit', scope: 'organization' });
  });

  it('errors when no link is returned', async () => {
    mockHttp.mockResolvedValueOnce(httpOk({}));
    const res = await createShareLink('01X');
    expect(res.ok).toBe(false);
  });
});

describe('sendFileToChat (chat conversations)', () => {
  it('uploads to OneDrive, shares, and posts a native file attachment', async () => {
    mockRead.mockResolvedValueOnce(Buffer.from('hi'));
    mockHttp
      .mockResolvedValueOnce(httpOk({ id: '01X', name: 'report.pdf' }))                          // upload (PUT)
      .mockResolvedValueOnce(httpOk({ link: { webUrl: 'https://share/r' } }))                    // createShareLink
      .mockResolvedValueOnce(shareInfoResponse('https://t-my.sharepoint.com/personal/u/Documents/Microsoft%20Teams%20Chat%20Files/report.pdf')); // getShareFileInfo
    mockSend.mockResolvedValueOnce(ok({ messageId: 'm', timestamp: 1 }) as never);

    const res = await sendFileToChat('48:notes', '/tmp/report.pdf', 'see this');
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value).toMatchObject({ conversationId: '48:notes', fileName: 'report.pdf', webUrl: 'https://share/r', messageId: 'm' });

    const [conv, content, options] = mockSend.mock.calls[0] as [string, string, { files?: Record<string, unknown>[] }];
    expect(conv).toBe('48:notes');
    expect(content).toBe('see this'); // caption is the message text; the file rides in the files property
    expect(options.files).toHaveLength(1);
    expect(options.files![0]).toMatchObject({
      '@type': 'http://schema.skype.com/File',
      title: 'report.pdf',
      id: 'LIU-GUID',
      objectUrl: expect.stringContaining('report.pdf'),
    });
  });

  it('posts with an empty caption when none is given', async () => {
    mockRead.mockResolvedValueOnce(Buffer.from('hi'));
    mockHttp
      .mockResolvedValueOnce(httpOk({ id: '01X', name: 'report.pdf' }))
      .mockResolvedValueOnce(httpOk({ link: { webUrl: 'https://share/r' } }))
      .mockResolvedValueOnce(shareInfoResponse('https://t-my.sharepoint.com/personal/u/Documents/report.pdf'));
    mockSend.mockResolvedValueOnce(ok({ messageId: 'm', timestamp: 1 }) as never);

    const res = await sendFileToChat('48:notes', '/tmp/report.pdf');
    expect(res.ok).toBe(true);
    const [, content, options] = mockSend.mock.calls[0] as [string, string, { files?: unknown[] }];
    expect(content).toBe('');
    expect(options.files).toHaveLength(1);
  });

  it('returns the upload error when uploading fails', async () => {
    mockRead.mockRejectedValueOnce(new Error('ENOENT'));
    const res = await sendFileToChat('48:notes', '/tmp/missing.pdf');
    expect(res.ok).toBe(false);
    expect(mockSend).not.toHaveBeenCalled();
  });

  it('returns the share-link error when no link is produced', async () => {
    mockRead.mockResolvedValueOnce(Buffer.from('hi'));
    mockHttp
      .mockResolvedValueOnce(httpOk({ id: '01X', name: 'report.pdf' })) // upload
      .mockResolvedValueOnce(httpOk({})); // createLink -> no link
    const res = await sendFileToChat('48:notes', '/tmp/report.pdf');
    expect(res.ok).toBe(false);
    expect(mockSend).not.toHaveBeenCalled();
  });

  it('returns the send error when posting the message fails', async () => {
    mockRead.mockResolvedValueOnce(Buffer.from('hi'));
    mockHttp
      .mockResolvedValueOnce(httpOk({ id: '01X', name: 'report.pdf' }))
      .mockResolvedValueOnce(httpOk({ link: { webUrl: 'https://share/r' } }))
      .mockResolvedValueOnce(shareInfoResponse('https://t-my.sharepoint.com/personal/u/Documents/report.pdf'));
    mockSend.mockResolvedValueOnce({ ok: false, error: { code: 'API_ERROR' } } as never);
    const res = await sendFileToChat('48:notes', '/tmp/report.pdf');
    expect(res.ok).toBe(false);
  });
});

describe('sendFileToChat (channel conversations)', () => {
  it('uploads into the channel SharePoint library and posts a native attachment (no org link)', async () => {
    mockChannelInfo.mockResolvedValueOnce(ok({ groupId: 'GID', sharepointSiteUrl: 'https://t.sharepoint.com/teams/X' }) as never);
    mockRead.mockResolvedValueOnce(Buffer.from('hi'));
    mockHttp
      .mockResolvedValueOnce(httpOk({ id: 'folder', parentReference: { driveId: 'drv' } }))        // getChannelFilesFolder
      .mockResolvedValueOnce(httpOk({ id: 'item1', name: 'report.pdf', webUrl: 'x' }))              // uploadFileToDriveFolder (PUT)
      .mockResolvedValueOnce(shareInfoResponse('https://t.sharepoint.com/teams/X/Shared%20Documents/Chan/report.pdf')); // getShareFileInfo(driveId)
    mockSend.mockResolvedValueOnce(ok({ messageId: 'm', timestamp: 1 }) as never);

    const res = await sendFileToChat('19:abc@thread.tacv2', '/tmp/report.pdf', 'cap');
    expect(res.ok).toBe(true);
    // folder lookup + upload + share-info = 3 calls; channels skip the org share link.
    expect(mockHttp).toHaveBeenCalledTimes(3);
    expect(mockHttp.mock.calls[2][0]).toContain('/drives/drv/items/');

    const [conv, content, options] = mockSend.mock.calls[0] as [string, string, { files?: Record<string, unknown>[] }];
    expect(conv).toBe('19:abc@thread.tacv2');
    expect(content).toBe('cap');
    expect(options.files![0].objectUrl).toContain('/teams/X/');
  });

  it('retries the channel upload under a de-duplicated name when the target is locked (HTTP 423)', async () => {
    mockChannelInfo.mockResolvedValueOnce(ok({ groupId: 'GID', sharepointSiteUrl: 'https://t.sharepoint.com/teams/X' }) as never);
    mockRead.mockResolvedValueOnce(Buffer.from('hi'));
    mockHttp
      .mockResolvedValueOnce(httpOk({ id: 'folder', parentReference: { driveId: 'drv' } }))          // getChannelFilesFolder
      .mockResolvedValueOnce(httpErr('HTTP 423: {"error":{"code":"resourceLocked"}}'))               // upload PUT (locked)
      .mockResolvedValueOnce(httpOk({ id: 'item1', name: 'report (2).pdf', webUrl: 'x' }))            // retry PUT succeeds
      .mockResolvedValueOnce(httpOk({                                                                 // getShareFileInfo (reflects the renamed file)
        id: 'item1', name: 'report (2).pdf',
        webUrl: 'https://t.sharepoint.com/teams/X/Shared%20Documents/Chan/report%20(2).pdf',
        sharepointIds: { listItemUniqueId: 'LIU-GUID', siteId: 'SITE-GUID' },
      }));
    mockSend.mockResolvedValueOnce(ok({ messageId: 'm', timestamp: 1 }) as never);

    const res = await sendFileToChat('19:abc@thread.tacv2', '/tmp/report.pdf', 'cap');
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    // folder lookup + locked PUT + retry PUT + share-info = 4 calls.
    expect(mockHttp).toHaveBeenCalledTimes(4);
    expect(mockHttp.mock.calls[1][0]).toContain('/drives/drv/items/folder:/report.pdf:/content');
    expect(mockHttp.mock.calls[2][0]).toContain('/drives/drv/items/folder:/report%20(2).pdf:/content');
    // The de-duplicated name propagates to the returned result and the chiclet.
    expect(res.value.fileName).toBe('report (2).pdf');
    const [, , options] = mockSend.mock.calls[0] as [string, string, { files?: Record<string, unknown>[] }];
    expect(options.files![0]).toMatchObject({ title: 'report (2).pdf', fileName: 'report (2).pdf' });
  });

  it('returns the channel-info error when the team cannot be resolved', async () => {
    mockChannelInfo.mockResolvedValueOnce({ ok: false, error: { code: 'API_ERROR' } } as never);
    const res = await sendFileToChat('19:abc@thread.tacv2', '/tmp/report.pdf');
    expect(res.ok).toBe(false);
    expect(mockSend).not.toHaveBeenCalled();
  });
});

describe('sendFilesToChat (multiple files, one message)', () => {
  it('uploads every file and posts them as attachments on a single message', async () => {
    mockRead.mockResolvedValue(Buffer.from('hi'));
    // Two files: each needs upload (PUT) + createShareLink + getShareFileInfo.
    mockHttp
      .mockResolvedValueOnce(httpOk({ id: '01A', name: 'a.pdf' }))
      .mockResolvedValueOnce(httpOk({ link: { webUrl: 'https://share/a' } }))
      .mockResolvedValueOnce(shareInfoResponse('https://t-my.sharepoint.com/personal/u/Documents/Microsoft%20Teams%20Chat%20Files/a.pdf'))
      .mockResolvedValueOnce(httpOk({ id: '01B', name: 'b.xlsx' }))
      .mockResolvedValueOnce(httpOk({ link: { webUrl: 'https://share/b' } }))
      .mockResolvedValueOnce(shareInfoResponse('https://t-my.sharepoint.com/personal/u/Documents/Microsoft%20Teams%20Chat%20Files/b.xlsx'));
    mockSend.mockResolvedValueOnce(ok({ messageId: 'm', timestamp: 1 }) as never);

    const res = await sendFilesToChat('48:notes', ['/tmp/a.pdf', '/tmp/b.xlsx'], 'two files');
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.files).toHaveLength(2);
    expect(res.value.messageId).toBe('m');

    // One single message carrying both file properties.
    expect(mockSend).toHaveBeenCalledTimes(1);
    const [conv, content, options] = mockSend.mock.calls[0] as [string, string, { files?: Record<string, unknown>[] }];
    expect(conv).toBe('48:notes');
    expect(content).toBe('two files');
    expect(options.files).toHaveLength(2);
  });

  it('rejects an empty file list and never posts', async () => {
    const res = await sendFilesToChat('48:notes', []);
    expect(res.ok).toBe(false);
    expect(mockSend).not.toHaveBeenCalled();
  });

  it('returns the first upload error and never posts', async () => {
    mockRead.mockRejectedValueOnce(new Error('ENOENT'));
    const res = await sendFilesToChat('48:notes', ['/tmp/missing.pdf', '/tmp/b.xlsx']);
    expect(res.ok).toBe(false);
    expect(mockSend).not.toHaveBeenCalled();
  });
});

describe('auth and error propagation', () => {
  const AUTH_ERR = { ok: false, error: { code: 'AUTH_REQUIRED' } } as never;
  const httpErr = { ok: false, error: { code: 'API_ERROR' } } as never;

  it('listDriveFiles returns the auth error and propagates http errors and empty values', async () => {
    mockAuth.mockReturnValueOnce(AUTH_ERR);
    expect((await listDriveFiles()).ok).toBe(false);
    expect(mockHttp).not.toHaveBeenCalled();

    mockHttp.mockResolvedValueOnce(httpErr);
    expect((await listDriveFiles()).ok).toBe(false);

    mockHttp.mockResolvedValueOnce(httpOk({})); // no value field -> defaults to []
    const res = await listDriveFiles();
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value.items).toEqual([]);
  });

  it('uploadFile returns the auth error, rejects oversize files and propagates http errors', async () => {
    mockAuth.mockReturnValueOnce(AUTH_ERR);
    expect((await uploadFile('/tmp/x.txt')).ok).toBe(false);

    mockRead.mockResolvedValueOnce({ length: 251 * 1024 * 1024 } as never);
    const big = await uploadFile('/tmp/big.bin');
    expect(big.ok).toBe(false);
    expect(mockHttp).not.toHaveBeenCalled();

    mockRead.mockResolvedValueOnce(Buffer.from('hi'));
    mockHttp.mockResolvedValueOnce(httpErr);
    expect((await uploadFile('/tmp/x.txt')).ok).toBe(false);
  });

  it('downloadFile returns the auth error and handles a text() rejection', async () => {
    mockAuth.mockReturnValueOnce(AUTH_ERR);
    expect((await downloadFile('01X', '/tmp/o')).ok).toBe(false);

    const fetchMock = vi.fn().mockResolvedValue({ ok: false, status: 500, text: async () => { throw new Error('no body'); } });
    vi.stubGlobal('fetch', fetchMock);
    const res = await downloadFile('01X', '/tmp/o');
    expect(res.ok).toBe(false);
    vi.unstubAllGlobals();
  });

  it('createShareLink returns the auth error and propagates http errors', async () => {
    mockAuth.mockReturnValueOnce(AUTH_ERR);
    expect((await createShareLink('01X')).ok).toBe(false);

    mockHttp.mockResolvedValueOnce(httpErr);
    expect((await createShareLink('01X')).ok).toBe(false);
  });
});
