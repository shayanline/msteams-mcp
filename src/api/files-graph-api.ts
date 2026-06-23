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
import { sendMessage, getChannelFilesInfo } from './chatsvc-messaging.js';
import { getConversationType } from '../utils/parsers.js';

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

/**
 * Downloads a file shared in a Teams conversation by its SharePoint/OneDrive URL.
 *
 * Works for files from other users' drives that you have access to (e.g. shared
 * in a chat). Uses the Microsoft Graph Shares API which can resolve any sharing
 * URL the authenticated user has been granted access to.
 */
export async function downloadSharedFile(shareUrl: string, outputPath: string): Promise<Result<{ outputPath: string; bytes: number }>> {
  const auth = requireGraphAuth();
  if (!auth.ok) return auth;

  // Graph Shares API requires the URL encoded as base64url with a 'u!' prefix.
  const encodedUrl = 'u!' + Buffer.from(shareUrl).toString('base64').replace(/=/g, '').replace(/\+/g, '-').replace(/\//g, '_');

  const res = await fetch(`${GRAPH_BASE_URL}/shares/${encodeURIComponent(encodedUrl)}/driveItem/content`, {
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

/** SharePoint identifiers needed to build a native Teams file chiclet. */
interface ShareFileInfo {
  /** SharePoint listItemUniqueId GUID (the chiclet's file id). */
  itemId: string;
  fileName: string;
  /** Lowercase extension without the dot, e.g. "zip", "docx". */
  fileType: string;
  /** Absolute SharePoint URL to the file (webUrl). */
  objectUrl: string;
  /** SharePoint site collection base URL, with a trailing slash. */
  baseUrl: string;
  /** SharePoint site GUID (sharepointIds.siteId). */
  siteId: string;
}

/**
 * Derives the SharePoint site base URL (with trailing slash) for a file. Teams
 * keys the chiclet's baseUrl on the site root, not the file path. Prefers the
 * site URL reported in sharepointIds, otherwise trims the file's webUrl at the
 * document library (OneDrive uses "Documents", team sites use "Shared
 * Documents").
 */
function deriveSiteBaseUrl(webUrl: string, sharepointSiteUrl?: string): string {
  if (sharepointSiteUrl) {
    return sharepointSiteUrl.endsWith('/') ? sharepointSiteUrl : `${sharepointSiteUrl}/`;
  }
  for (const marker of ['/Shared%20Documents/', '/Shared Documents/', '/Documents/']) {
    const idx = webUrl.indexOf(marker);
    if (idx > 0) return webUrl.slice(0, idx + 1);
  }
  return '';
}

/**
 * Fetches the SharePoint identifiers for an uploaded drive item that Teams needs
 * to render a native file chiclet. `sharepointIds.listItemUniqueId` is the id
 * Teams keys the file on, and `webUrl` is the openable SharePoint URL. Pass a
 * `driveId` for items in a team/channel library (otherwise the user's drive).
 */
async function getShareFileInfo(driveItemId: string, driveId?: string): Promise<Result<ShareFileInfo>> {
  const auth = requireGraphAuth();
  if (!auth.ok) return auth;

  const itemPath = driveId
    ? `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(driveItemId)}`
    : `/me/drive/items/${encodeURIComponent(driveItemId)}`;

  const response = await httpRequest<Record<string, unknown>>(
    `${GRAPH_BASE_URL}${itemPath}?$select=id,name,webUrl,sharepointIds`,
    { method: 'GET', headers: { ...bearer(auth.value), Accept: 'application/json' } }
  );
  if (!response.ok) return response;

  const item = response.value.data;
  const sp = (item.sharepointIds ?? {}) as Record<string, unknown>;
  const name = (item.name as string) ?? 'file';
  const dot = name.lastIndexOf('.');
  const fileType = dot > 0 ? name.slice(dot + 1).toLowerCase() : '';
  const objectUrl = (item.webUrl as string) || '';

  return ok({
    itemId: (sp.listItemUniqueId as string) || (item.id as string),
    fileName: name,
    fileType,
    objectUrl,
    baseUrl: deriveSiteBaseUrl(objectUrl, sp.siteUrl as string | undefined),
    siteId: (sp.siteId as string) || '',
  });
}

/** A channel's files folder location in its team SharePoint library. */
interface ChannelFilesFolder {
  driveId: string;
  folderId: string;
}

/**
 * Resolves a channel's files folder (drive id + folder item id) via Graph, so a
 * file can be uploaded into the channel's own SharePoint library rather than the
 * sender's OneDrive.
 */
async function getChannelFilesFolder(groupId: string, channelId: string): Promise<Result<ChannelFilesFolder>> {
  const auth = requireGraphAuth();
  if (!auth.ok) return auth;

  const response = await httpRequest<Record<string, unknown>>(
    `${GRAPH_BASE_URL}/teams/${encodeURIComponent(groupId)}/channels/${encodeURIComponent(channelId)}/filesFolder?$select=id,parentReference`,
    { method: 'GET', headers: { ...bearer(auth.value), Accept: 'application/json' } }
  );
  if (!response.ok) return response;

  const folder = response.value.data;
  const parent = (folder.parentReference ?? {}) as Record<string, unknown>;
  const driveId = parent.driveId as string | undefined;
  const folderId = folder.id as string | undefined;
  if (!driveId || !folderId) {
    return err(createError(ErrorCode.API_ERROR, "Could not resolve the channel's files folder."));
  }
  return ok({ driveId, folderId });
}

/** Uploads a local file into a specific drive folder (used for channel libraries). */
async function uploadFileToDriveFolder(
  driveId: string,
  folderId: string,
  localPath: string
): Promise<Result<DriveItem>> {
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
  const response = await httpRequest<Record<string, unknown>>(
    `${GRAPH_BASE_URL}/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(folderId)}:/${encodeURIComponent(name)}:/content`,
    { method: 'PUT', headers: { ...bearer(auth.value), 'Content-Type': 'application/octet-stream' }, body: new Uint8Array(data) }
  );
  if (!response.ok) return response;
  return ok(parseItem(response.value.data));
}

/**
 * Builds a chatsvc file object (`http://schema.skype.com/File`) describing a
 * SharePoint-hosted file. Serialised into the message's `files` property, this
 * is what makes Teams render a native file chiclet. The shape mirrors what the
 * Teams web client posts for a shared document (image-only fields like
 * amsreferences/filePreview are omitted).
 */
function buildFileProperty(info: ShareFileInfo, shareUrl: string): Record<string, unknown> {
  return {
    '@type': 'http://schema.skype.com/File',
    version: 2,
    id: info.itemId,
    baseUrl: info.baseUrl,
    type: info.fileType,
    title: info.fileName,
    state: 'active',
    objectUrl: info.objectUrl,
    providerData: '',
    itemid: info.itemId,
    fileName: info.fileName,
    fileType: info.fileType,
    fileInfo: {
      itemId: null,
      fileUrl: info.objectUrl,
      siteUrl: info.baseUrl,
      serverRelativeUrl: '',
      shareUrl,
      shareId: '',
    },
    chicletBreadcrumbs: null,
    botFileProperties: {},
    isUploadError: null,
    progressComplete: null,
    permissionScope: 'users',
    filePreview: null,
    sharepointIds: {
      listId: null,
      listItemUniqueId: info.itemId,
      siteId: info.siteId || null,
      siteUrl: null,
      webId: null,
    },
    publication: null,
    site: null,
    fileChicletState: { serviceName: 'p2p', state: 'active' },
  };
}

/** One uploaded file ready to be attached to a message. */
interface PreparedFile {
  fileProperty: Record<string, unknown>;
  fileName: string;
  webUrl: string;
}

/**
 * Uploads a single local file to the right place for the conversation and
 * returns the chatsvc file property plus display info, without posting anything.
 * Channel files go into the channel's SharePoint library; chat files go into the
 * sender's OneDrive and are shared via an org link.
 */
async function prepareFileForChat(
  conversationId: string,
  localPath: string
): Promise<Result<PreparedFile>> {
  const isChannel = getConversationType(conversationId) === 'channel';

  let info: Result<ShareFileInfo>;
  let webUrl: string;
  let shareUrl: string;

  if (isChannel) {
    // Upload into the channel's SharePoint library; members already have access.
    const channel = await getChannelFilesInfo(conversationId);
    if (!channel.ok) return channel;
    const folder = await getChannelFilesFolder(channel.value.groupId, conversationId);
    if (!folder.ok) return folder;
    const uploaded = await uploadFileToDriveFolder(folder.value.driveId, folder.value.folderId, localPath);
    if (!uploaded.ok) return uploaded;

    info = await getShareFileInfo(uploaded.value.id, folder.value.driveId);
    webUrl = uploaded.value.webUrl ?? '';
    shareUrl = '';
  } else {
    // Upload to the user's OneDrive chat-files area and grant access via an org link.
    const uploaded = await uploadFile(localPath, 'Microsoft Teams Chat Files');
    if (!uploaded.ok) return uploaded;
    const link = await createShareLink(uploaded.value.id, 'edit');
    if (!link.ok) return link;

    info = await getShareFileInfo(uploaded.value.id);
    webUrl = link.value.webUrl;
    shareUrl = link.value.webUrl;
  }

  if (!info.ok) return info;

  return ok({
    fileProperty: buildFileProperty(info.value, shareUrl),
    fileName: info.value.fileName,
    webUrl,
  });
}

/**
 * Sends a file into a Teams conversation as a native attachment: it posts a
 * message carrying the file in its `files` property so Teams shows a real file
 * chiclet (and lists it in the Files tab), with an optional caption as the text.
 *
 * The file is uploaded to the correct place for the conversation type, matching
 * what the Teams client itself does:
 * - **Channels**: the channel's own SharePoint files folder (so it appears in the
 *   channel Files tab); channel members already have access.
 * - **Chats** (1:1, group, meeting, self): the sender's OneDrive "Microsoft Teams
 *   Chat Files", shared with the conversation via an org link.
 */
export async function sendFileToChat(
  conversationId: string,
  localPath: string,
  caption?: string
): Promise<Result<{ conversationId: string; fileName: string; webUrl: string; messageId: string }>> {
  const prepared = await prepareFileForChat(conversationId, localPath);
  if (!prepared.ok) return prepared;

  const sent = await sendMessage(conversationId, caption ?? '', { files: [prepared.value.fileProperty] });
  if (!sent.ok) return sent;

  return ok({ conversationId, fileName: prepared.value.fileName, webUrl: prepared.value.webUrl, messageId: sent.value.messageId });
}

/**
 * Sends several local files into a Teams conversation as native attachments on a
 * single message (one message, multiple file chiclets), with an optional caption
 * as the text. Each file is uploaded to the right place for the conversation type
 * (see `sendFileToChat`), then all are attached to one `sendMessage` call.
 */
export async function sendFilesToChat(
  conversationId: string,
  localPaths: string[],
  caption?: string
): Promise<Result<{ conversationId: string; files: Array<{ fileName: string; webUrl: string }>; messageId: string }>> {
  if (localPaths.length === 0) {
    return err(createError(ErrorCode.INVALID_INPUT, 'Provide at least one file to send.'));
  }

  const prepared: PreparedFile[] = [];
  for (const localPath of localPaths) {
    const file = await prepareFileForChat(conversationId, localPath);
    if (!file.ok) return file;
    prepared.push(file.value);
  }

  const sent = await sendMessage(conversationId, caption ?? '', {
    files: prepared.map((p) => p.fileProperty),
  });
  if (!sent.ok) return sent;

  return ok({
    conversationId,
    files: prepared.map((p) => ({ fileName: p.fileName, webUrl: p.webUrl })),
    messageId: sent.value.messageId,
  });
}
