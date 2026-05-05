import type { DriveLocation } from './drive-location.js';
import { DEFAULT_DRIVE_LOCATION } from './drive-location.js';
import {
  callGraphAt,
  driveItemPath,
  GraphApiError,
  type GraphResponse,
  getGraphBaseUrl,
  graphError,
  graphResult
} from './graph-client.js';

export interface DriveItemPreviewResponse {
  getUrl?: string;
  postParameters?: unknown;
  /** Present when preview uses POST URL flow */
  postUrl?: string;
}

/**
 * POST …/drive/items/{id}/preview — embeddable preview session for Office documents (Word, PowerPoint, Excel, …).
 * @see https://learn.microsoft.com/en-us/graph/api/driveitem-preview
 */
export async function createDriveItemPreview(
  token: string,
  itemId: string,
  body: Record<string, unknown>,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION,
  graphBaseUrl: string = getGraphBaseUrl()
): Promise<GraphResponse<DriveItemPreviewResponse>> {
  try {
    const path = `${driveItemPath(location, itemId)}/preview`;
    const r = await callGraphAt<DriveItemPreviewResponse>(graphBaseUrl, token, path, {
      method: 'POST',
      body: JSON.stringify(body)
    });
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to create preview', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to create preview');
  }
}
