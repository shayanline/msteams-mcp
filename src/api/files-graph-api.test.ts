import { describe, it, expect, vi, beforeEach } from 'vitest';
import { ok } from '../types/result.js';

vi.mock('../utils/http.js', () => ({ httpRequest: vi.fn() }));
vi.mock('../utils/auth-guards.js', () => ({ requireGraphAuth: vi.fn() }));
vi.mock('./chatsvc-messaging.js', () => ({ sendMessage: vi.fn() }));
vi.mock('node:fs/promises', () => ({ readFile: vi.fn(), writeFile: vi.fn() }));

import { httpRequest } from '../utils/http.js';
import { requireGraphAuth } from '../utils/auth-guards.js';
import { sendMessage } from './chatsvc-messaging.js';
import { readFile, writeFile } from 'node:fs/promises';
import { listDriveFiles, uploadFile, downloadFile, createShareLink, sendFileToChat } from './files-graph-api.js';

const mockHttp = vi.mocked(httpRequest);
const mockAuth = vi.mocked(requireGraphAuth);
const mockSend = vi.mocked(sendMessage);
const mockRead = vi.mocked(readFile);
const mockWrite = vi.mocked(writeFile);
const httpOk = (data: unknown) => ok({ status: 200, headers: new Headers(), data } as never);

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

describe('sendFileToChat', () => {
  it('uploads, shares, and posts a message linking the file', async () => {
    mockRead.mockResolvedValueOnce(Buffer.from('hi'));
    mockHttp
      .mockResolvedValueOnce(httpOk({ id: '01X', name: 'report.pdf' }))      // upload
      .mockResolvedValueOnce(httpOk({ link: { webUrl: 'https://share/r' } })); // createLink
    mockSend.mockResolvedValueOnce(ok({ messageId: 'm', timestamp: 1 }) as never);
    const res = await sendFileToChat('48:notes', '/tmp/report.pdf', 'see this');
    expect(res.ok).toBe(true);
    if (!res.ok) return;
    expect(res.value).toMatchObject({ conversationId: '48:notes', fileName: 'report.pdf', webUrl: 'https://share/r' });
    const [conv, content] = mockSend.mock.calls[0] as [string, string];
    expect(conv).toBe('48:notes');
    expect(content).toContain('[report.pdf](https://share/r)');
    expect(content).toContain('see this');
  });

  it('posts without a caption when none is given', async () => {
    mockRead.mockResolvedValueOnce(Buffer.from('hi'));
    mockHttp
      .mockResolvedValueOnce(httpOk({ id: '01X', name: 'report.pdf' }))
      .mockResolvedValueOnce(httpOk({ link: { webUrl: 'https://share/r' } }));
    mockSend.mockResolvedValueOnce(ok({ messageId: 'm', timestamp: 1 }) as never);
    const res = await sendFileToChat('48:notes', '/tmp/report.pdf');
    expect(res.ok).toBe(true);
    const [, content] = mockSend.mock.calls[0] as [string, string];
    expect(content).toBe('[report.pdf](https://share/r)');
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
      .mockResolvedValueOnce(httpOk({ link: { webUrl: 'https://share/r' } }));
    mockSend.mockResolvedValueOnce({ ok: false, error: { code: 'API_ERROR' } } as never);
    const res = await sendFileToChat('48:notes', '/tmp/report.pdf');
    expect(res.ok).toBe(false);
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
