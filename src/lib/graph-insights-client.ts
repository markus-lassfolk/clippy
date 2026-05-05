import {
  callGraph,
  type DriveLocation,
  driveItemPath,
  GraphApiError,
  type GraphResponse,
  graphError,
  graphErrorFromApiError,
  graphResult
} from './graph-client.js';
import { graphUserPath } from './graph-user-path.js';

/**
 * Office Graph [insights](https://learn.microsoft.com/graph/api/resources/officegraphinsights):
 * `trending` (documents trending around the user), `used` (recently used by the user),
 * `shared` (shared with the user). All three are **delegated** under `/me/insights/...`
 * (or `/users/{id}/insights/...` for another user when the caller has consent).
 */
export type InsightKind = 'trending' | 'used' | 'shared';

/** Minimal shape covering the common surface across the three insight types. */
export interface InsightItem {
  id?: string;
  weight?: number;
  /** trending */
  resourceVisualization?: {
    title?: string;
    type?: string;
    mediaType?: string;
    containerWebUrl?: string;
    containerDisplayName?: string;
    previewImageUrl?: string;
  };
  resourceReference?: {
    webUrl?: string;
    id?: string;
    type?: string;
  };
  /** used */
  lastUsed?: { lastAccessedDateTime?: string; lastModifiedDateTime?: string };
  /** shared */
  lastShared?: {
    sharedDateTime?: string;
    sharingReference?: { webUrl?: string };
    sharingSubject?: string;
    sharingType?: string;
    sharedBy?: { displayName?: string; address?: string };
  };
}

export interface InsightListResponse {
  value?: InsightItem[];
  '@odata.nextLink'?: string;
}

export async function listInsights(
  token: string,
  kind: InsightKind,
  options: { user?: string; top?: number } = {}
): Promise<GraphResponse<InsightListResponse>> {
  const top = options.top && options.top > 0 ? `?$top=${Math.min(Math.max(1, options.top), 200)}` : '';
  const path = `${graphUserPath(options.user, `insights/${kind}`)}${top}`;
  try {
    return await callGraph<InsightListResponse>(token, path);
  } catch (err) {
    if (err instanceof GraphApiError) return graphErrorFromApiError(err);
    return graphError(err instanceof Error ? err.message : `Failed to list insights/${kind}`);
  }
}

export interface ItemActivity {
  id?: string;
  action?: Record<string, unknown>;
  actor?: { user?: { displayName?: string; email?: string }; application?: { displayName?: string } };
  times?: { recordedDateTime?: string; observedDateTime?: string };
}

export interface ItemActivitiesResponse {
  value?: ItemActivity[];
  '@odata.nextLink'?: string;
}

/** `GET /drives/{driveId}/items/{itemId}/activities` — per-item activity feed. */
export async function listDriveItemActivities(
  token: string,
  loc: DriveLocation,
  itemId: string,
  options: { top?: number } = {}
): Promise<GraphResponse<ItemActivitiesResponse>> {
  const top = options.top && options.top > 0 ? `?$top=${Math.min(Math.max(1, options.top), 200)}` : '';
  const path = `${driveItemPath(loc, itemId)}/activities${top}`;
  try {
    return await callGraph<ItemActivitiesResponse>(token, path);
  } catch (err) {
    if (err instanceof GraphApiError) return graphErrorFromApiError(err);
    return graphError(err instanceof Error ? err.message : 'Failed to list drive item activities');
  }
}

export interface DriveItemPreview {
  getUrl?: string;
  postUrl?: string;
  postParameters?: string;
}

/**
 * `POST /drives/{driveId}/items/{itemId}/preview` — preview session URL for any drive item.
 * Body fields are optional ([driveItem-preview](https://learn.microsoft.com/graph/api/driveitem-preview)).
 */
export async function createDriveItemPreview(
  token: string,
  loc: DriveLocation,
  itemId: string,
  body: { page?: number | string; zoom?: number; allowEdit?: boolean; chromeless?: boolean } = {}
): Promise<GraphResponse<DriveItemPreview>> {
  const path = `${driveItemPath(loc, itemId)}/preview`;
  const cleaned: Record<string, unknown> = {};
  if (body.page !== undefined) cleaned.page = body.page;
  if (body.zoom !== undefined) cleaned.zoom = body.zoom;
  if (body.allowEdit !== undefined) cleaned.allowEdit = body.allowEdit;
  if (body.chromeless !== undefined) cleaned.chromeless = body.chromeless;
  try {
    const r = await callGraph<DriveItemPreview>(token, path, {
      method: 'POST',
      body: JSON.stringify(cleaned)
    });
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to create preview', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphErrorFromApiError(err);
    return graphError(err instanceof Error ? err.message : 'Failed to create preview');
  }
}

export interface FollowedSite {
  id?: string;
  webUrl?: string;
  displayName?: string;
  name?: string;
  description?: string;
}

export interface FollowedSitesResponse {
  value?: FollowedSite[];
  '@odata.nextLink'?: string;
}

/** `GET /me/followedSites` — sites the user follows. */
export async function listFollowedSites(token: string): Promise<GraphResponse<FollowedSitesResponse>> {
  try {
    return await callGraph<FollowedSitesResponse>(token, '/me/followedSites');
  } catch (err) {
    if (err instanceof GraphApiError) return graphErrorFromApiError(err);
    return graphError(err instanceof Error ? err.message : 'Failed to list /me/followedSites');
  }
}

/** `POST /me/followedSites/add` — follow one or more sites. */
export async function followSites(token: string, siteIds: string[]): Promise<GraphResponse<FollowedSitesResponse>> {
  if (siteIds.length === 0) {
    return graphError('Provide at least one site id');
  }
  const body = { value: siteIds.map((id) => ({ id })) };
  try {
    return await callGraph<FollowedSitesResponse>(token, '/me/followedSites/add', {
      method: 'POST',
      body: JSON.stringify(body)
    });
  } catch (err) {
    if (err instanceof GraphApiError) return graphErrorFromApiError(err);
    return graphError(err instanceof Error ? err.message : 'Failed to follow site');
  }
}

/** `POST /me/followedSites/remove` — unfollow one or more sites. */
export async function unfollowSites(token: string, siteIds: string[]): Promise<GraphResponse<void>> {
  if (siteIds.length === 0) {
    return graphError('Provide at least one site id');
  }
  const body = { value: siteIds.map((id) => ({ id })) };
  try {
    return await callGraph<void>(
      token,
      '/me/followedSites/remove',
      { method: 'POST', body: JSON.stringify(body) },
      false
    );
  } catch (err) {
    if (err instanceof GraphApiError) return graphErrorFromApiError(err);
    return graphError(err instanceof Error ? err.message : 'Failed to unfollow site');
  }
}
