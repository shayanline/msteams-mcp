/**
 * Unit tests for file tools (real handlers).
 */
import { describe, it, expect, vi, beforeEach } from 'vitest';
import { ok, err } from '../types/result.js';
import { createError, ErrorCode } from '../types/errors.js';
// Initialise the tool barrel first so the registry's eager tool-array spreads
// resolve before any single tool module is imported (avoids a circular-init crash).
import './index.js';

vi.mock('../api/files-api.js', () => ({ getSharedFiles: vi.fn() }));
vi.mock('../api/files-graph-api.js', () => ({
  listDriveFiles: vi.fn(),
  uploadFile: vi.fn(),
  downloadFile: vi.fn(),
  sendFileToChat: vi.fn(),
  sendFilesToChat: vi.fn(),
}));

import { getSharedFiles } from '../api/files-api.js';
import { listDriveFiles, uploadFile, downloadFile, sendFileToChat, sendFilesToChat } from '../api/files-graph-api.js';
import {
  getSharedFilesTool,
  listFilesTool,
  uploadFileTool,
  downloadFileTool,
  sendFileTool,
  sendFilesTool,
} from './file-tools.js';

const ctx = { server: {} } as never;
const anErr = err(createError(ErrorCode.API_ERROR, 'boom'));

beforeEach(() => vi.clearAllMocks());

describe('listFilesTool', () => {
  it('lists files on success', async () => {
    vi.mocked(listDriveFiles).mockResolvedValue(ok({ items: [{ id: 'i1' }] }) as never);
    const res = await listFilesTool.handler({}, ctx);
    expect(res.success).toBe(true);
    if (res.success) expect(res.data).toEqual({ count: 1, items: [{ id: 'i1' }] });
  });
  it('propagates errors', async () => {
    vi.mocked(listDriveFiles).mockResolvedValue(anErr as never);
    expect((await listFilesTool.handler({}, ctx)).success).toBe(false);
  });
});

describe('uploadFileTool', () => {
  it('uploads on success', async () => {
    vi.mocked(uploadFile).mockResolvedValue(ok({ id: 'i1' }) as never);
    const res = await uploadFileTool.handler({ localPath: '/tmp/a' }, ctx);
    expect(res.success).toBe(true);
    if (res.success) expect(res.data.item).toEqual({ id: 'i1' });
  });
  it('propagates errors', async () => {
    vi.mocked(uploadFile).mockResolvedValue(anErr as never);
    expect((await uploadFileTool.handler({ localPath: '/tmp/a' }, ctx)).success).toBe(false);
  });
});

describe('downloadFileTool', () => {
  it('downloads on success', async () => {
    vi.mocked(downloadFile).mockResolvedValue(ok({ path: '/tmp/a' }) as never);
    const res = await downloadFileTool.handler({ itemId: 'i1', outputPath: '/tmp/a' }, ctx);
    expect(res.success).toBe(true);
    if (res.success) expect(res.data.message).toBe('File downloaded.');
  });
  it('propagates errors', async () => {
    vi.mocked(downloadFile).mockResolvedValue(anErr as never);
    expect((await downloadFileTool.handler({ itemId: 'i1', outputPath: '/tmp/a' }, ctx)).success).toBe(false);
  });
});

describe('sendFileTool', () => {
  it('sends on success', async () => {
    vi.mocked(sendFileToChat).mockResolvedValue(ok({ messageId: 'm1' }) as never);
    const res = await sendFileTool.handler({ conversationId: 'c1', localPath: '/tmp/a' }, ctx);
    expect(res.success).toBe(true);
    if (res.success) expect(res.data.message).toBe('File sent to conversation.');
  });
  it('propagates errors', async () => {
    vi.mocked(sendFileToChat).mockResolvedValue(anErr as never);
    expect((await sendFileTool.handler({ conversationId: 'c1', localPath: '/tmp/a' }, ctx)).success).toBe(false);
  });
});

describe('sendFilesTool', () => {
  it('sends multiple files on one message on success', async () => {
    vi.mocked(sendFilesToChat).mockResolvedValue(ok({ messageId: 'm1', files: [{}, {}] }) as never);
    const res = await sendFilesTool.handler({ conversationId: 'c1', localPaths: ['/tmp/a', '/tmp/b'] }, ctx);
    expect(res.success).toBe(true);
    if (res.success) expect(res.data.message).toBe('2 file(s) sent on one message.');
  });
  it('propagates errors', async () => {
    vi.mocked(sendFilesToChat).mockResolvedValue(anErr as never);
    expect((await sendFilesTool.handler({ conversationId: 'c1', localPaths: ['/tmp/a'] }, ctx)).success).toBe(false);
  });
});

describe('getSharedFilesTool', () => {
  it('returns files with hasMore when skipToken present', async () => {
    vi.mocked(getSharedFiles).mockResolvedValue(
      ok({ conversationId: 'c1', returned: 1, files: [{}], skipToken: 'tok' }) as never
    );
    const res = await getSharedFilesTool.handler({ conversationId: 'c1', pageSize: 25 }, ctx);
    expect(res.success).toBe(true);
    if (res.success) expect(res.data).toMatchObject({ hasMore: true, skipToken: 'tok' });
  });
  it('returns files with hasMore false when no skipToken', async () => {
    vi.mocked(getSharedFiles).mockResolvedValue(
      ok({ conversationId: 'c1', returned: 1, files: [{}] }) as never
    );
    const res = await getSharedFilesTool.handler({ conversationId: 'c1', pageSize: 25 }, ctx);
    expect(res.success).toBe(true);
    if (res.success) expect(res.data.hasMore).toBe(false);
  });
  it('propagates errors', async () => {
    vi.mocked(getSharedFiles).mockResolvedValue(anErr as never);
    expect((await getSharedFilesTool.handler({ conversationId: 'c1', pageSize: 25 }, ctx)).success).toBe(false);
  });
});
