import { callGraph, GraphApiError, fetchAllPages, type GraphResponse, graphError, graphResult } from './graph-client.js';
import { graphUserPath } from './graph-user-path.js';

/** User or group returned as directoryObject for manager / directReports (common fields). */
export interface OrgDirectoryObject {
  id: string;
  displayName?: string;
  mail?: string;
  userPrincipalName?: string;
  jobTitle?: string;
  '@odata.type'?: string;
}

/**
 * Get the manager of the signed-in user or of `forUser` (GET /me/manager, GET /users/{id}/manager).
 * Typically needs **User.Read** (self) or **User.Read.All** / directory scopes for other users.
 */
export async function getManager(token: string, forUser?: string): Promise<GraphResponse<OrgDirectoryObject>> {
  try {
    const path = graphUserPath(forUser, 'manager');
    const r = await callGraph<OrgDirectoryObject>(token, path);
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to get manager', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get manager');
  }
}

/**
 * List direct reports for /me or another user (GET …/directReports). Paginated.
 */
export async function listDirectReports(
  token: string,
  forUser?: string
): Promise<GraphResponse<OrgDirectoryObject[]>> {
  const path = graphUserPath(forUser, 'directReports');
  return fetchAllPages<OrgDirectoryObject>(token, path, 'Failed to list direct reports');
}
