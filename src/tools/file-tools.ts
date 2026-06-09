/**
 * File-related tool handlers.
 */

import { z } from 'zod';
import type { Tool } from '@modelcontextprotocol/sdk/types.js';
import type { RegisteredTool, ToolContext, ToolResult } from './index.js';
import { handleApiResult } from './index.js';
import { getSharedFiles } from '../api/files-api.js';
import { listDriveFiles, uploadFile, downloadFile, sendFileToChat } from '../api/files-graph-api.js';
import {
  DEFAULT_FILES_PAGE_SIZE,
  MAX_FILES_PAGE_SIZE,
} from '../constants.js';

// ─────────────────────────────────────────────────────────────────────────────
// Schemas
// ─────────────────────────────────────────────────────────────────────────────

export const GetSharedFilesInputSchema = z.object({
  conversationId: z.string().min(1),
  pageSize: z.number().min(1).max(MAX_FILES_PAGE_SIZE).optional().default(DEFAULT_FILES_PAGE_SIZE),
  skipToken: z.string().optional(),
});

export const ListFilesInputSchema = z.object({
  folderPath: z.string().optional(),
});

export const UploadFileInputSchema = z.object({
  localPath: z.string().min(1),
  folder: z.string().optional(),
});

export const DownloadFileInputSchema = z.object({
  itemId: z.string().min(1),
  outputPath: z.string().min(1),
});

export const SendFileInputSchema = z.object({
  conversationId: z.string().min(1),
  localPath: z.string().min(1),
  caption: z.string().optional(),
});

// ─────────────────────────────────────────────────────────────────────────────
// Tool Definitions
// ─────────────────────────────────────────────────────────────────────────────

const getSharedFilesToolDefinition: Tool = {
  name: 'teams_get_shared_files',
  description: 'Get files and links shared in a Teams conversation. Returns file names, URLs, extensions, sizes, and who shared them. Works for channels, group chats, 1:1 chats, and meeting chats. Use the conversationId from other tools (teams_get_favorites, teams_search, teams_find_channel, teams_get_chat). Supports pagination via skipToken for conversations with many files.',
  inputSchema: {
    type: 'object',
    properties: {
      conversationId: {
        type: 'string',
        description: 'The conversation ID to get shared files for (e.g., "19:abc@thread.tacv2" for a channel, or a chat conversation ID).',
      },
      pageSize: {
        type: 'number',
        description: `Number of items per page (default: ${DEFAULT_FILES_PAGE_SIZE}, max: ${MAX_FILES_PAGE_SIZE})`,
      },
      skipToken: {
        type: 'string',
        description: 'Continuation token from a previous response to get the next page of results.',
      },
    },
    required: ['conversationId'],
  },
};

const listFilesToolDefinition: Tool = {
  name: 'teams_list_files',
  description: 'List files and folders in the user\'s OneDrive. Pass a folderPath (e.g. "Documents/Reports") to browse into a folder, or omit it for the root. Returns item ids (for download), names, sizes and web URLs.',
  inputSchema: {
    type: 'object',
    properties: {
      folderPath: { type: 'string', description: 'OneDrive folder path to list (omit for root).' },
    },
  },
};

const uploadFileToolDefinition: Tool = {
  name: 'teams_upload_file',
  description: 'Upload a local file to the user\'s OneDrive (up to 250 MB). Returns the drive item id and web URL. To share it into a chat, use teams_send_file instead.',
  inputSchema: {
    type: 'object',
    properties: {
      localPath: { type: 'string', description: 'Absolute path to the local file to upload.' },
      folder: { type: 'string', description: 'OneDrive folder to upload into (default Apps/AutomationUploads).' },
    },
    required: ['localPath'],
  },
};

const downloadFileToolDefinition: Tool = {
  name: 'teams_download_file',
  description: 'Download a OneDrive file by its item id (from teams_list_files) to a local path.',
  inputSchema: {
    type: 'object',
    properties: {
      itemId: { type: 'string', description: 'OneDrive drive item id.' },
      outputPath: { type: 'string', description: 'Local path to write the downloaded file to.' },
    },
    required: ['itemId', 'outputPath'],
  },
};

const sendFileToolDefinition: Tool = {
  name: 'teams_send_file',
  description: 'Send a local file into a Teams conversation as a native file attachment. This is the correct way to share a file: it posts a real file chiclet (icon, preview, open in Teams) that also appears in the conversation\'s Files tab, not just a link. The file is uploaded to the right place automatically: a channel\'s own SharePoint files folder for channels, or your OneDrive "Microsoft Teams Chat Files" (shared via an org link) for 1:1, group, meeting and self chats. Works for all conversation types. Add an optional caption as the message text. Confirm the content with the user before sending.',
  inputSchema: {
    type: 'object',
    properties: {
      conversationId: { type: 'string', description: 'The conversation to send the file to (1:1, group, meeting, self chat, or channel).' },
      localPath: { type: 'string', description: 'Absolute path to the local file to send.' },
      caption: { type: 'string', description: 'Optional message text to send with the file.' },
    },
    required: ['conversationId', 'localPath'],
  },
};

// ─────────────────────────────────────────────────────────────────────────────
// Handlers
// ─────────────────────────────────────────────────────────────────────────────

async function handleListFiles(
  input: z.infer<typeof ListFilesInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  return handleApiResult(await listDriveFiles(input.folderPath), (v) => ({ count: v.items.length, items: v.items }));
}

async function handleUploadFile(
  input: z.infer<typeof UploadFileInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  return handleApiResult(await uploadFile(input.localPath, input.folder), (v) => ({ item: v, message: 'File uploaded to OneDrive.' }));
}

async function handleDownloadFile(
  input: z.infer<typeof DownloadFileInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  return handleApiResult(await downloadFile(input.itemId, input.outputPath), (v) => ({ ...v, message: 'File downloaded.' }));
}

async function handleSendFile(
  input: z.infer<typeof SendFileInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  return handleApiResult(await sendFileToChat(input.conversationId, input.localPath, input.caption), (v) => ({ ...v, message: 'File sent to conversation.' }));
}

async function handleGetSharedFiles(
  input: z.infer<typeof GetSharedFilesInputSchema>,
  _ctx: ToolContext
): Promise<ToolResult> {
  const result = await getSharedFiles(input.conversationId, {
    pageSize: input.pageSize,
    skipToken: input.skipToken,
  });

  return handleApiResult(result, (value) => ({
    conversationId: value.conversationId,
    returned: value.returned,
    files: value.files,
    ...(value.skipToken ? { skipToken: value.skipToken, hasMore: true } : { hasMore: false }),
  }));
}

// ─────────────────────────────────────────────────────────────────────────────
// Exports
// ─────────────────────────────────────────────────────────────────────────────

export const getSharedFilesTool: RegisteredTool<typeof GetSharedFilesInputSchema> = {
  definition: getSharedFilesToolDefinition,
  schema: GetSharedFilesInputSchema,
  handler: handleGetSharedFiles,
};

export const listFilesTool: RegisteredTool<typeof ListFilesInputSchema> = {
  definition: listFilesToolDefinition, schema: ListFilesInputSchema, handler: handleListFiles,
};
export const uploadFileTool: RegisteredTool<typeof UploadFileInputSchema> = {
  definition: uploadFileToolDefinition, schema: UploadFileInputSchema, handler: handleUploadFile,
};
export const downloadFileTool: RegisteredTool<typeof DownloadFileInputSchema> = {
  definition: downloadFileToolDefinition, schema: DownloadFileInputSchema, handler: handleDownloadFile,
};
export const sendFileTool: RegisteredTool<typeof SendFileInputSchema> = {
  definition: sendFileToolDefinition, schema: SendFileInputSchema, handler: handleSendFile,
};

/** All file-related tools. */
export const fileTools = [getSharedFilesTool, listFilesTool, uploadFileTool, downloadFileTool, sendFileTool];
