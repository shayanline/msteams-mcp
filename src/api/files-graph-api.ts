/**
 * OneDrive / SharePoint files via Microsoft Graph (Files.ReadWrite.All).
 *
 * Upload, list, download files, and share a file into a Teams chat by uploading
 * it to the user's OneDrive, creating an organisation share link, and posting
 * that link as a message (Teams unfurls it into a file card).
 */

import { readFile, writeFile } from 'node:fs/promises';
import { basename } from 'node:path';
import { httpRequest } from '../utils/http.js';
import { GRAPH_BASE_URL } from '../utils/api-config.js';
import { type Result, ok, err } from '../types/result.js';
import { ErrorCode, createError } from '../types/errors.js';
import { requireGraphAuth } from '../utils/auth-guards.js';
import { sendMessage } from './chatsvc-messaging.js';

const GRAPH_UPLOAD_MAX = 250 * 1024 * 1024; // simple upload supports up to 250 MB

export interface DriveItem {
  id: string;
  name: string;
  webUrl?: string;
  size?: number;
  isFolder: boolean;
}

function bearer(token: string): Record<string, string> {
  return { Authorization: `Bearer ${token}` };
}

function parseItem(raw: Record<string, unknown>): DriveItem {
  return {
    id: raw.id as string,
    name: raw.name as string,
    webUrl: raw.webUrl as string | undefined,
    size: raw.size as number | undefined,
    isFolder: raw.folder !== undefined,
  };
}

/** Lists items in a OneDrive folder (root by default). */
export async function listDriveFiles(folderPath?: string): Promise<Result<{ items: DriveItem[] }>> {
  const auth = requireGraphAuth();
  if (!auth.ok) return auth;

  const path = folderPath
    ? `/me/drive/root:/${encodeURIComponent(folderPath).replace(/%2F/g, '/')}:/children`
    : '/me/drive/root/children';

  const response = await httpRequest<{ value: Array<Record<string, unknown>> }>(
    `${GRAPH_BASE_URL}${path}?$top=100&$select=id,name,webUrl,size,folder`,
    { method: 'GET', headers: { ...bearer(auth.value), Accept: 'application/json' } }
  );
  if (!response.ok) return response;
  return ok({ items: (response.value.data.value ?? []).map(parseItem) });
}

/** Uploads a local file to OneDrive under the given folder. Returns the drive item. */
export async function uploadFile(localPath: string, folder = 'Apps/AutomationUploads'): Promise<Result<DriveItem>> {
  const auth = requireGraphAuth();
  if (!auth.ok) return auth;

  let data: Buffer;
  try {
    data = await readFile(localPath);
  } catch {
    return err(createError(ErrorCode.INVALID_INPUT, `Could not read file: ${localPath}`));
  }
  if (data.length > GRAPH_UPLOAD_MAX) {
    return err(createError(ErrorCode.INVALID_INPUT, 'File exceeds the 250 MB upload limit.'));
  }

  const name = basename(localPath);
  const drivePath = `${folder}/${name}`.split('/').map(encodeURIComponent).join('/');
  const response = await httpRequest<Record<string, unknown>>(
    `${GRAPH_BASE_URL}/me/drive/root:/${drivePath}:/content`,
    { method: 'PUT', headers: { ...bearer(auth.value), 'Content-Type': 'application/octet-stream' }, body: new Uint8Array(data) }
  );
  if (!response.ok) return response;
  return ok(parseItem(response.value.data));
}

/** Downloads a OneDrive file (by item id) to a local path. */
export async function downloadFile(itemId: string, outputPath: string): Promise<Result<{ outputPath: string; bytes: number }>> {
  const auth = requireGraphAuth();
  if (!auth.ok) return auth;

  // Binary download, so use a raw fetch rather than the JSON-oriented httpRequest.
  const res = await fetch(`${GRAPH_BASE_URL}/me/drive/items/${encodeURIComponent(itemId)}/content`, {
    method: 'GET',
    headers: bearer(auth.value),
  });
  if (!res.ok) {
    const text = await res.text().catch(() => '');
    return err(createError(ErrorCode.UNKNOWN, `Download failed: HTTP ${res.status} ${text.slice(0, 150)}`));
  }
  const buf = Buffer.from(await res.arrayBuffer());
  await writeFile(outputPath, buf);
  return ok({ outputPath, bytes: buf.length });
}

/** Creates an organisation-scoped sharing link for a drive item. */
export async function createShareLink(
  itemId: string,
  type: 'view' | 'edit' = 'view'
): Promise<Result<{ webUrl: string }>> {
  const auth = requireGraphAuth();
  if (!auth.ok) return auth;

  const response = await httpRequest<Record<string, unknown>>(
    `${GRAPH_BASE_URL}/me/drive/items/${encodeURIComponent(itemId)}/createLink`,
    {
      method: 'POST',
      headers: { ...bearer(auth.value), 'Content-Type': 'application/json' },
      body: JSON.stringify({ type, scope: 'organization' }),
    }
  );
  if (!response.ok) return response;
  const link = response.value.data.link as { webUrl?: string } | undefined;
  if (!link?.webUrl) return err(createError(ErrorCode.UNKNOWN, 'Share link was not returned.'));
  return ok({ webUrl: link.webUrl });
}

/**
 * Sends a file into a Teams conversation: uploads it to OneDrive (Teams chat
 * files area), creates an org edit link, and posts a message with the file
 * linked by name plus an optional caption. Teams renders the SharePoint link as
 * a file card.
 */
export async function sendFileToChat(
  conversationId: string,
  localPath: string,
  caption?: string
): Promise<Result<{ conversationId: string; fileName: string; webUrl: string }>> {
  const uploaded = await uploadFile(localPath, 'Microsoft Teams Chat Files');
  if (!uploaded.ok) return uploaded;

  const link = await createShareLink(uploaded.value.id, 'edit');
  if (!link.ok) return link;

  const intro = caption ? `${caption}\n\n` : '';
  const content = `${intro}[${uploaded.value.name}](${link.value.webUrl})`;
  const sent = await sendMessage(conversationId, content);
  if (!sent.ok) return sent;

  return ok({ conversationId, fileName: uploaded.value.name, webUrl: link.value.webUrl });
}
