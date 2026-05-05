import {
  callGraph,
  fetchAllPages,
  GraphApiError,
  type GraphResponse,
  getGraphBaseUrl,
  graphError,
  graphResult
} from './graph-client.js';
import { graphUserPath } from './graph-user-path.js';

const USER_SELECT =
  'id,displayName,givenName,surname,mail,userPrincipalName,jobTitle,department,officeLocation,mobilePhone,businessPhones';

/** User or group returned as directoryObject for manager / directReports (common fields). */
export interface OrgDirectoryObject {
  id: string;
  displayName?: string;
  mail?: string;
  userPrincipalName?: string;
  jobTitle?: string;
  '@odata.type'?: string;
}

/** Directory user profile (subset of user resource). */
export interface OrgUserProfile {
  id: string;
  displayName?: string;
  givenName?: string;
  surname?: string;
  mail?: string;
  userPrincipalName?: string;
  jobTitle?: string;
  department?: string;
  officeLocation?: string;
  mobilePhone?: string;
  businessPhones?: string[];
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
export async function listDirectReports(token: string, forUser?: string): Promise<GraphResponse<OrgDirectoryObject[]>> {
  const path = graphUserPath(forUser, 'directReports');
  return fetchAllPages<OrgDirectoryObject>(token, path, 'Failed to list direct reports');
}

/**
 * Full management chain below a user (GET …/transitiveReports). Requires **User.Read** (self) or **User.Read.All** (others).
 */
export async function listTransitiveReports(
  token: string,
  forUser?: string
): Promise<GraphResponse<OrgDirectoryObject[]>> {
  const path = graphUserPath(forUser, 'transitiveReports');
  return fetchAllPages<OrgDirectoryObject>(token, path, 'Failed to list transitive reports', getGraphBaseUrl(), {
    headers: { ConsistencyLevel: 'eventual' }
  });
}

/**
 * GET /me or GET /users/{id} with a practical `$select` for directory display.
 */
export async function getUserProfile(token: string, userId?: string): Promise<GraphResponse<OrgUserProfile>> {
  const path = userId?.trim()
    ? `/users/${encodeURIComponent(userId.trim())}?$select=${USER_SELECT}`
    : `/me?$select=${USER_SELECT}`;
  try {
    const r = await callGraph<OrgUserProfile>(token, path);
    if (!r.ok || !r.data) {
      return graphError(r.error?.message || 'Failed to get user', r.error?.code, r.error?.status);
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get user');
  }
}
