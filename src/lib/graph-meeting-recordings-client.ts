import { createWriteStream } from 'node:fs';
import { mkdir, unlink } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import {
  callGraph,
  callGraphAbsolute,
  fetchGraphRaw,
  GraphApiError,
  type GraphResponse,
  graphError,
  graphErrorFromApiError,
  graphResult
} from './graph-client.js';
import { graphUserPath } from './graph-user-path.js';

/** `microsoft.graph.callRecording` (v1.0/beta). */
export interface CallRecording {
  id: string;
  meetingId?: string;
  callId?: string;
  recordingContentUrl?: string;
  contentCorrelationId?: string;
  createdDateTime?: string;
  endDateTime?: string;
  meetingOrganizer?: unknown;
}

/** `microsoft.graph.callTranscript` (v1.0/beta). */
export interface CallTranscript {
  id: string;
  meetingId?: string;
  callId?: string;
  transcriptContentUrl?: string;
  contentCorrelationId?: string;
  createdDateTime?: string;
  endDateTime?: string;
  meetingOrganizer?: unknown;
}

export interface CallRecordingListResponse {
  value?: CallRecording[];
  '@odata.nextLink'?: string;
  '@odata.deltaLink'?: string;
}

export interface CallTranscriptListResponse {
  value?: CallTranscript[];
  '@odata.nextLink'?: string;
  '@odata.deltaLink'?: string;
}

function meetingsRoot(user?: string): string {
  return graphUserPath(user, 'onlineMeetings');
}

/** `GET /me/onlineMeetings/{id}/recordings` — recordings on a single meeting. */
export async function listMeetingRecordings(
  token: string,
  meetingId: string,
  user?: string
): Promise<GraphResponse<CallRecordingListResponse>> {
  const path = `${meetingsRoot(user)}/${encodeURIComponent(meetingId)}/recordings`;
  try {
    return await callGraph<CallRecordingListResponse>(token, path);
  } catch (err) {
    if (err instanceof GraphApiError) return graphErrorFromApiError(err);
    return graphError(err instanceof Error ? err.message : 'Failed to list recordings');
  }
}

/** `GET /me/onlineMeetings/{id}/transcripts` — transcripts on a single meeting. */
export async function listMeetingTranscripts(
  token: string,
  meetingId: string,
  user?: string
): Promise<GraphResponse<CallTranscriptListResponse>> {
  const path = `${meetingsRoot(user)}/${encodeURIComponent(meetingId)}/transcripts`;
  try {
    return await callGraph<CallTranscriptListResponse>(token, path);
  } catch (err) {
    if (err instanceof GraphApiError) return graphErrorFromApiError(err);
    return graphError(err instanceof Error ? err.message : 'Failed to list transcripts');
  }
}

/** `GET /me/onlineMeetings/{id}/recordings/{recordingId}/content` — binary stream. */
export function recordingContentPath(meetingId: string, recordingId: string, user?: string): string {
  return `${meetingsRoot(user)}/${encodeURIComponent(meetingId)}/recordings/${encodeURIComponent(recordingId)}/content`;
}

/** `GET /me/onlineMeetings/{id}/transcripts/{transcriptId}/content` — VTT stream. */
export function transcriptContentPath(meetingId: string, transcriptId: string, user?: string): string {
  return `${meetingsRoot(user)}/${encodeURIComponent(meetingId)}/transcripts/${encodeURIComponent(transcriptId)}/content`;
}

/** `GET .../transcripts/{id}/metadataContent` — time-aligned utterance metadata. */
export function transcriptMetadataContentPath(meetingId: string, transcriptId: string, user?: string): string {
  return `${meetingsRoot(user)}/${encodeURIComponent(meetingId)}/transcripts/${encodeURIComponent(transcriptId)}/metadataContent`;
}

/**
 * `getAllRecordings(meetingOrganizerUserId='@id',startDateTime=@start,endDateTime=@end)` —
 * tenant-wide (or per-organizer) recording roll-up. Pass `pageUrl` to follow `@odata.nextLink`.
 */
export async function getAllRecordings(
  token: string,
  args: { organizerUserId: string; start: string; end: string; user?: string; pageUrl?: string; top?: number }
): Promise<GraphResponse<CallRecordingListResponse>> {
  if (args.pageUrl?.trim()) {
    try {
      return await callGraphAbsolute<CallRecordingListResponse>(token, args.pageUrl.trim());
    } catch (err) {
      if (err instanceof GraphApiError) return graphErrorFromApiError(err);
      return graphError(err instanceof Error ? err.message : 'Failed to follow recordings page');
    }
  }
  const fn =
    `getAllRecordings(meetingOrganizerUserId='${encodeURIComponent(args.organizerUserId)}'` +
    `,startDateTime=${encodeURIComponent(args.start)}` +
    `,endDateTime=${encodeURIComponent(args.end)})`;
  let path = `${meetingsRoot(args.user)}/${fn}`;
  if (args.top !== undefined && Number.isFinite(args.top) && args.top > 0) {
    path += `?$top=${Math.min(999, Math.floor(args.top))}`;
  }
  try {
    return await callGraph<CallRecordingListResponse>(token, path);
  } catch (err) {
    if (err instanceof GraphApiError) return graphErrorFromApiError(err);
    return graphError(err instanceof Error ? err.message : 'Failed to call getAllRecordings');
  }
}

/**
 * `getAllTranscripts(...)` — same shape as `getAllRecordings`, returns `callTranscript` list.
 */
export async function getAllTranscripts(
  token: string,
  args: { organizerUserId: string; start: string; end: string; user?: string; pageUrl?: string; top?: number }
): Promise<GraphResponse<CallTranscriptListResponse>> {
  if (args.pageUrl?.trim()) {
    try {
      return await callGraphAbsolute<CallTranscriptListResponse>(token, args.pageUrl.trim());
    } catch (err) {
      if (err instanceof GraphApiError) return graphErrorFromApiError(err);
      return graphError(err instanceof Error ? err.message : 'Failed to follow transcripts page');
    }
  }
  const fn =
    `getAllTranscripts(meetingOrganizerUserId='${encodeURIComponent(args.organizerUserId)}'` +
    `,startDateTime=${encodeURIComponent(args.start)}` +
    `,endDateTime=${encodeURIComponent(args.end)})`;
  let path = `${meetingsRoot(args.user)}/${fn}`;
  if (args.top !== undefined && Number.isFinite(args.top) && args.top > 0) {
    path += `?$top=${Math.min(999, Math.floor(args.top))}`;
  }
  try {
    return await callGraph<CallTranscriptListResponse>(token, path);
  } catch (err) {
    if (err instanceof GraphApiError) return graphErrorFromApiError(err);
    return graphError(err instanceof Error ? err.message : 'Failed to call getAllTranscripts');
  }
}

const MAX_DOWNLOAD_BYTES = 5 * 1024 * 1024 * 1024;
const ALLOWED_DOWNLOAD_HOSTS = [
  'graph.microsoft.com',
  'graph.microsoft.us',
  'dod-graph.microsoft.us',
  'microsoftgraph.chinacloudapi.cn',
  'graph.microsoft.de',
  'sharepoint.com',
  'sharepoint.us',
  'sharepoint.cn',
  'onedrive.live.com',
  'files.1drv.com',
  'stream.microsoft.com'
];

function isAllowedDownloadHost(hostname: string): boolean {
  return ALLOWED_DOWNLOAD_HOSTS.some((host) => hostname === host || hostname.endsWith(`.${host}`));
}

async function streamResponseToFile(body: ReadableStream<Uint8Array>, filePath: string): Promise<number> {
  const stream = createWriteStream(filePath, { flags: 'w', mode: 0o600 });
  let bytesWritten = 0;
  try {
    for await (const chunk of body) {
      if (bytesWritten + chunk.byteLength > MAX_DOWNLOAD_BYTES) {
        stream.destroy();
        await unlink(filePath).catch(() => {});
        throw new Error(`Download exceeded ${MAX_DOWNLOAD_BYTES} bytes`);
      }
      // codeql[js/http-to-file-access]: Writes authenticated Graph/stream media bytes to user-chosen output with host allowlist and size cap.
      if (!stream.write(chunk)) {
        await new Promise<void>((res, rej) => {
          stream.once('drain', res);
          stream.once('error', rej);
        });
      }
      bytesWritten += chunk.byteLength;
    }
    await new Promise<void>((res, rej) => stream.end((e?: Error | null) => (e ? rej(e) : res())));
    return bytesWritten;
  } catch (err) {
    stream.destroy();
    await unlink(filePath).catch(() => {});
    throw err;
  }
}

/**
 * Download a Graph media stream (recording or transcript content) to disk.
 * Follows a single redirect into the allowlisted CDN/Stream hosts.
 */
export async function downloadMediaToFile(
  token: string,
  graphPath: string,
  outputPath: string
): Promise<GraphResponse<{ path: string; bytes: number }>> {
  const target = resolve(outputPath);
  await mkdir(dirname(target), { recursive: true });
  try {
    const initial = await fetchGraphRaw(token, graphPath, { redirect: 'manual' });
    let response: Response = initial;
    if (initial.status >= 300 && initial.status < 400) {
      const location = initial.headers.get('location');
      if (!location) return graphError('Missing redirect location');
      let url: URL;
      try {
        url = new URL(location);
      } catch {
        return graphError('Redirect location is not a valid URL');
      }
      if (url.protocol !== 'https:') return graphError('Only HTTPS redirects are permitted');
      if (!isAllowedDownloadHost(url.hostname)) {
        return graphError(`Redirect host '${url.hostname}' is not allowlisted`);
      }
      response = await fetch(url.toString(), { redirect: 'manual' });
    }
    if (!response.ok) {
      const txt = await response.text().catch(() => '');
      return graphError(`Download failed: HTTP ${response.status}${txt ? ` ${txt.slice(0, 256)}` : ''}`);
    }
    if (!response.body) return graphError('Empty response body');
    const bytes = await streamResponseToFile(response.body, target);
    return graphResult({ path: target, bytes });
  } catch (err) {
    if (err instanceof GraphApiError) return graphErrorFromApiError(err);
    return graphError(err instanceof Error ? err.message : 'Download failed');
  }
}

/** Initial-window fields for the first page of `getAllRecordings(...)/delta`. */
export interface MeetingRecordingsDeltaInitial {
  organizerUserId: string;
  startDateTime: string;
  endDateTime: string;
  top?: number;
}

/** `getAllRecordings(...)/delta` — incremental sync (not per-meeting `.../recordings/delta`). */
export async function getRecordingsDeltaPage(
  token: string,
  args: {
    pageUrl?: string;
    user?: string;
    initial?: MeetingRecordingsDeltaInitial;
  }
): Promise<GraphResponse<CallRecordingListResponse>> {
  const pageUrl = args.pageUrl?.trim();
  if (pageUrl) {
    try {
      return await callGraphAbsolute<CallRecordingListResponse>(token, pageUrl);
    } catch (err) {
      if (err instanceof GraphApiError) return graphErrorFromApiError(err);
      return graphError(err instanceof Error ? err.message : 'Failed to follow recordings delta page');
    }
  }
  const init = args.initial;
  if (!init?.organizerUserId?.trim() || !init.startDateTime?.trim() || !init.endDateTime?.trim()) {
    return graphError(
      'Recordings delta needs a saved next/delta URL, or organizer plus start/end for the initial `getAllRecordings(...)/delta` request.'
    );
  }
  const fn =
    `getAllRecordings(meetingOrganizerUserId='${encodeURIComponent(init.organizerUserId.trim())}'` +
    `,startDateTime=${encodeURIComponent(init.startDateTime)}` +
    `,endDateTime=${encodeURIComponent(init.endDateTime)})`;
  let path = `${meetingsRoot(args.user)}/${fn}/delta`;
  if (init.top !== undefined && Number.isFinite(init.top) && init.top > 0) {
    path += `?$top=${Math.min(999, Math.floor(init.top))}`;
  }
  try {
    return await callGraph<CallRecordingListResponse>(token, path);
  } catch (err) {
    if (err instanceof GraphApiError) return graphErrorFromApiError(err);
    return graphError(err instanceof Error ? err.message : 'Failed to call getAllRecordings delta');
  }
}

/** Initial-window fields for the first page of `getAllTranscripts(...)/delta`. */
export interface MeetingTranscriptsDeltaInitial {
  organizerUserId: string;
  startDateTime: string;
  endDateTime: string;
  top?: number;
}

/** `getAllTranscripts(...)/delta` — incremental sync (not per-meeting `.../transcripts/delta`). */
export async function getTranscriptsDeltaPage(
  token: string,
  args: {
    pageUrl?: string;
    user?: string;
    initial?: MeetingTranscriptsDeltaInitial;
  }
): Promise<GraphResponse<CallTranscriptListResponse>> {
  const pageUrl = args.pageUrl?.trim();
  if (pageUrl) {
    try {
      return await callGraphAbsolute<CallTranscriptListResponse>(token, pageUrl);
    } catch (err) {
      if (err instanceof GraphApiError) return graphErrorFromApiError(err);
      return graphError(err instanceof Error ? err.message : 'Failed to follow transcripts delta page');
    }
  }
  const init = args.initial;
  if (!init?.organizerUserId?.trim() || !init.startDateTime?.trim() || !init.endDateTime?.trim()) {
    return graphError(
      'Transcripts delta needs a saved next/delta URL, or organizer plus start/end for the initial `getAllTranscripts(...)/delta` request.'
    );
  }
  const fn =
    `getAllTranscripts(meetingOrganizerUserId='${encodeURIComponent(init.organizerUserId.trim())}'` +
    `,startDateTime=${encodeURIComponent(init.startDateTime)}` +
    `,endDateTime=${encodeURIComponent(init.endDateTime)})`;
  let path = `${meetingsRoot(args.user)}/${fn}/delta`;
  if (init.top !== undefined && Number.isFinite(init.top) && init.top > 0) {
    path += `?$top=${Math.min(999, Math.floor(init.top))}`;
  }
  try {
    return await callGraph<CallTranscriptListResponse>(token, path);
  } catch (err) {
    if (err instanceof GraphApiError) return graphErrorFromApiError(err);
    return graphError(err instanceof Error ? err.message : 'Failed to call getAllTranscripts delta');
  }
}
