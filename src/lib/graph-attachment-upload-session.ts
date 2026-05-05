/**
 * Large file attachments for Outlook mail and calendar events via Graph upload session
 * (POST …/attachments/createUploadSession + chunked PUT to uploadUrl).
 */

import { callGraph, GraphApiError, type GraphResponse, graphError, graphResult } from './graph-client.js';
import { graphUserPath } from './graph-user-path.js';

/** Raw file size above which we prefer upload session over inline base64 POST. */
export const GRAPH_OUTLOOK_ATTACHMENT_SESSION_THRESHOLD_BYTES = 2 * 1024 * 1024;

export interface GraphAttachmentUploadSession {
  uploadUrl: string;
  expirationDateTime?: string;
}

const CHUNK_SIZE = 4 * 1024 * 1024;

/**
 * Upload bytes to a pre-authorized Graph upload URL (no Bearer on PUT).
 * Returns parsed JSON from the final successful response when present.
 */
export async function uploadBufferViaGraphUploadUrl(
  uploadUrl: string,
  data: Buffer
): Promise<GraphResponse<Record<string, unknown> | undefined>> {
  const total = data.byteLength;
  if (total === 0) {
    return graphError('Cannot upload zero-byte attachment via upload session', undefined, 400);
  }
  let start = 0;
  let lastJson: Record<string, unknown> | undefined;
  while (start < total) {
    const end = Math.min(start + CHUNK_SIZE, total);
    const slice = data.subarray(start, end);
    const contentRange = `bytes ${start}-${end - 1}/${total}`;
    let response: Response;
    try {
      // codeql[js/file-access-to-http]: Chunked PUT to Graph-provided uploadUrl; body is the caller's attachment bytes, not arbitrary file exfiltration.
      response = await fetch(uploadUrl, {
        method: 'PUT',
        headers: {
          'Content-Length': String(slice.byteLength),
          'Content-Range': contentRange
        },
        body: new Blob([Uint8Array.from(slice)])
      });
    } catch (err) {
      return graphError(err instanceof Error ? err.message : 'Upload chunk failed');
    }
    const text = await response.text();
    if (!response.ok) {
      return graphError(text || `Upload failed: HTTP ${response.status}`, undefined, response.status);
    }
    if (text) {
      try {
        const parsed = JSON.parse(text) as Record<string, unknown>;
        if (parsed && typeof parsed === 'object') {
          lastJson = parsed;
        }
      } catch {
        // non-JSON success body
      }
    }
    start = end;
  }
  return graphResult(lastJson);
}

export async function createMailMessageFileAttachmentUploadSession(
  token: string,
  messageId: string,
  name: string,
  size: number,
  contentType: string,
  user?: string
): Promise<GraphResponse<GraphAttachmentUploadSession>> {
  const path = `${graphUserPath(user, `messages/${encodeURIComponent(messageId)}/attachments/createUploadSession`)}`;
  const body = {
    AttachmentItem: {
      attachmentType: 'file',
      name,
      size,
      isInline: false,
      contentType: contentType || 'application/octet-stream'
    }
  };
  try {
    const result = await callGraph<GraphAttachmentUploadSession>(token, path, {
      method: 'POST',
      body: JSON.stringify(body)
    });
    if (!result.ok || !result.data?.uploadUrl) {
      return graphError(
        result.error?.message || 'Failed to create message attachment upload session',
        result.error?.code,
        result.error?.status
      );
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to create message attachment upload session');
  }
}

export async function createCalendarEventFileAttachmentUploadSession(
  token: string,
  eventId: string,
  name: string,
  size: number,
  contentType: string,
  user?: string
): Promise<GraphResponse<GraphAttachmentUploadSession>> {
  const path = `${graphUserPath(user, `events/${encodeURIComponent(eventId)}/attachments/createUploadSession`)}`;
  const body = {
    AttachmentItem: {
      attachmentType: 'file',
      name,
      size,
      isInline: false,
      contentType: contentType || 'application/octet-stream'
    }
  };
  try {
    const result = await callGraph<GraphAttachmentUploadSession>(token, path, {
      method: 'POST',
      body: JSON.stringify(body)
    });
    if (!result.ok || !result.data?.uploadUrl) {
      return graphError(
        result.error?.message || 'Failed to create event attachment upload session',
        result.error?.code,
        result.error?.status
      );
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to create event attachment upload session');
  }
}
