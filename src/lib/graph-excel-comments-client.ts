/**
 * Excel workbook threaded comments on a drive item (Microsoft Graph **beta**).
 * @see https://learn.microsoft.com/en-us/graph/api/resources/workbookcomment
 */
import type { DriveLocation } from './drive-location.js';
import { DEFAULT_DRIVE_LOCATION, driveRootPrefix } from './drive-location.js';
import {
  callGraphAt,
  fetchAllPages,
  GraphApiError,
  type GraphResponse,
  graphError,
  graphErrorFromApiError
} from './graph-client.js';
import { getGraphBetaUrl } from './graph-constants.js';

export type WorkbookCommentJson = Record<string, unknown>;

function commentsCollectionPath(location: DriveLocation, itemId: string): string {
  return `${driveRootPrefix(location)}/items/${encodeURIComponent(itemId.trim())}/workbook/comments`;
}

function commentItemPath(location: DriveLocation, itemId: string, commentId: string): string {
  return `${commentsCollectionPath(location, itemId)}/${encodeURIComponent(commentId.trim())}`;
}

export async function listExcelWorkbookComments(
  token: string,
  itemId: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION
): Promise<GraphResponse<WorkbookCommentJson[]>> {
  return fetchAllPages<WorkbookCommentJson>(
    token,
    commentsCollectionPath(location, itemId),
    'Failed to list workbook comments',
    getGraphBetaUrl()
  );
}

export async function getExcelWorkbookComment(
  token: string,
  itemId: string,
  commentId: string,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION
): Promise<GraphResponse<WorkbookCommentJson>> {
  try {
    return await callGraphAt<WorkbookCommentJson>(
      getGraphBetaUrl(),
      token,
      commentItemPath(location, itemId, commentId)
    );
  } catch (err) {
    if (err instanceof GraphApiError) return graphErrorFromApiError(err);
    return graphError(err instanceof Error ? err.message : 'Failed to get workbook comment');
  }
}

export async function createExcelWorkbookComment(
  token: string,
  itemId: string,
  body: Record<string, unknown>,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION
): Promise<GraphResponse<WorkbookCommentJson>> {
  try {
    return await callGraphAt<WorkbookCommentJson>(getGraphBetaUrl(), token, commentsCollectionPath(location, itemId), {
      method: 'POST',
      body: JSON.stringify(body)
    });
  } catch (err) {
    if (err instanceof GraphApiError) return graphErrorFromApiError(err);
    return graphError(err instanceof Error ? err.message : 'Failed to create workbook comment');
  }
}

export async function addExcelWorkbookCommentReply(
  token: string,
  itemId: string,
  commentId: string,
  body: Record<string, unknown>,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION
): Promise<GraphResponse<WorkbookCommentJson>> {
  try {
    const path = `${commentItemPath(location, itemId, commentId)}/replies`;
    return await callGraphAt<WorkbookCommentJson>(getGraphBetaUrl(), token, path, {
      method: 'POST',
      body: JSON.stringify(body)
    });
  } catch (err) {
    if (err instanceof GraphApiError) return graphErrorFromApiError(err);
    return graphError(err instanceof Error ? err.message : 'Failed to add comment reply');
  }
}

export async function patchExcelWorkbookComment(
  token: string,
  itemId: string,
  commentId: string,
  body: Record<string, unknown>,
  location: DriveLocation = DEFAULT_DRIVE_LOCATION
): Promise<GraphResponse<WorkbookCommentJson>> {
  try {
    return await callGraphAt<WorkbookCommentJson>(
      getGraphBetaUrl(),
      token,
      commentItemPath(location, itemId, commentId),
      {
        method: 'PATCH',
        body: JSON.stringify(body)
      }
    );
  } catch (err) {
    if (err instanceof GraphApiError) return graphErrorFromApiError(err);
    return graphError(err instanceof Error ? err.message : 'Failed to patch workbook comment');
  }
}
