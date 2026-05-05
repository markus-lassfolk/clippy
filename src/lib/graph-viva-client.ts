/**
 * Microsoft Graph **beta** surfaces used by Viva / work-time / insights-related flows.
 * Paths and delegated permission support vary by tenant; many operations are preview-only.
 */

import { callGraphAt, fetchAllPages, GraphApiError, type GraphResponse, graphError } from './graph-client.js';
import { GRAPH_BETA_URL } from './graph-constants.js';

function workingTimeSchedulePath(userId?: string): string {
  const u = userId?.trim();
  if (u) return `/users/${encodeURIComponent(u)}/solutions/workingTimeSchedule`;
  return '/me/solutions/workingTimeSchedule';
}

function itemInsightsPath(userId?: string): string {
  const u = userId?.trim();
  if (u) return `/users/${encodeURIComponent(u)}/settings/itemInsights`;
  return '/me/settings/itemInsights';
}

function assignedRolesPath(userId?: string): string {
  const u = userId?.trim();
  if (u) return `/users/${encodeURIComponent(u)}/employeeExperience/assignedRoles`;
  return '/me/employeeExperience/assignedRoles';
}

function employeeExperiencePath(userId?: string): string {
  const u = userId?.trim();
  if (u) return `/users/${encodeURIComponent(u)}/employeeExperience`;
  return '/me/employeeExperience';
}

function learningCourseActivitiesCollectionPath(userId?: string): string {
  const u = userId?.trim();
  if (u) return `/users/${encodeURIComponent(u)}/employeeExperience/learningCourseActivities`;
  return '/me/employeeExperience/learningCourseActivities';
}

function assignedRoleItemPath(userId: string | undefined, roleId: string): string {
  return `${assignedRolesPath(userId)}/${encodeURIComponent(roleId.trim())}`;
}

function assignedRoleMembersCollectionPath(userId: string | undefined, roleId: string): string {
  return `${assignedRoleItemPath(userId, roleId)}/members`;
}

function assignedRoleMemberItemPath(userId: string | undefined, roleId: string, memberId: string): string {
  return `${assignedRoleMembersCollectionPath(userId, roleId)}/${encodeURIComponent(memberId.trim())}`;
}

/** Escape a string for use inside OData single-quoted literals (e.g. alternate keys). */
export function escapeODataSingleQuotedKey(value: string): string {
  return value.replace(/'/g, "''");
}

/** OData query for Viva list endpoints (`$filter`, `$select`, `$top`, `$skip`, `$count`). */
export function buildVivaListQuery(opts: {
  filter?: string;
  select?: string;
  top?: number;
  skip?: number;
  count?: boolean;
}): string {
  const parts: string[] = [];
  if (opts.filter?.trim()) parts.push(`$filter=${encodeURIComponent(opts.filter.trim())}`);
  if (opts.select?.trim()) parts.push(`$select=${encodeURIComponent(opts.select.trim())}`);
  if (opts.top !== undefined && Number.isFinite(opts.top) && opts.top > 0) {
    parts.push(`$top=${encodeURIComponent(String(Math.floor(opts.top)))}`);
  }
  if (opts.skip !== undefined && Number.isFinite(opts.skip) && opts.skip > 0) {
    parts.push(`$skip=${encodeURIComponent(String(Math.floor(opts.skip)))}`);
  }
  if (opts.count) parts.push('$count=true');
  if (!parts.length) return '';
  return `?${parts.join('&')}`;
}

export async function getWorkingTimeSchedule(token: string, userId?: string): Promise<GraphResponse<unknown>> {
  try {
    return await callGraphAt<unknown>(GRAPH_BETA_URL, token, workingTimeSchedulePath(userId));
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get workingTimeSchedule');
  }
}

export async function patchWorkingTimeSchedule(
  token: string,
  body: unknown,
  userId?: string
): Promise<GraphResponse<unknown>> {
  try {
    return await callGraphAt<unknown>(GRAPH_BETA_URL, token, workingTimeSchedulePath(userId), {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    });
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to patch workingTimeSchedule');
  }
}

export async function deleteWorkingTimeSchedule(
  token: string,
  userId?: string,
  ifMatch?: string
): Promise<GraphResponse<void>> {
  const headers: Record<string, string> = {};
  if (ifMatch?.trim()) headers['If-Match'] = ifMatch.trim();
  try {
    return await callGraphAt<void>(
      GRAPH_BETA_URL,
      token,
      workingTimeSchedulePath(userId),
      { method: 'DELETE', headers },
      false
    );
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to delete workingTimeSchedule');
  }
}

export async function startWorkingTime(token: string, userId?: string): Promise<GraphResponse<void>> {
  const path = `${workingTimeSchedulePath(userId)}/startWorkingTime`;
  try {
    return await callGraphAt<void>(GRAPH_BETA_URL, token, path, { method: 'POST' }, false);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to startWorkingTime');
  }
}

export async function endWorkingTime(token: string, userId?: string): Promise<GraphResponse<void>> {
  const path = `${workingTimeSchedulePath(userId)}/endWorkingTime`;
  try {
    return await callGraphAt<void>(GRAPH_BETA_URL, token, path, { method: 'POST' }, false);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to endWorkingTime');
  }
}

export async function getUserItemInsightsSettings(token: string, userId?: string): Promise<GraphResponse<unknown>> {
  try {
    return await callGraphAt<unknown>(GRAPH_BETA_URL, token, itemInsightsPath(userId));
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get itemInsights settings');
  }
}

export async function patchUserItemInsightsSettings(
  token: string,
  body: unknown,
  userId?: string
): Promise<GraphResponse<unknown>> {
  try {
    return await callGraphAt<unknown>(GRAPH_BETA_URL, token, itemInsightsPath(userId), {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    });
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to patch itemInsights settings');
  }
}

export async function deleteUserItemInsightsSettings(
  token: string,
  userId?: string,
  ifMatch?: string
): Promise<GraphResponse<void>> {
  const headers: Record<string, string> = {};
  if (ifMatch?.trim()) headers['If-Match'] = ifMatch.trim();
  try {
    return await callGraphAt<void>(
      GRAPH_BETA_URL,
      token,
      itemInsightsPath(userId),
      { method: 'DELETE', headers },
      false
    );
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to delete itemInsights settings');
  }
}

export async function listEmployeeExperienceAssignedRoles(
  token: string,
  userId?: string,
  listQuery: string = ''
): Promise<GraphResponse<unknown[]>> {
  const q = listQuery.startsWith('?') ? listQuery : listQuery ? `?${listQuery}` : '';
  return fetchAllPages<unknown>(
    token,
    `${assignedRolesPath(userId)}${q}`,
    'Failed to list employeeExperience assignedRoles',
    GRAPH_BETA_URL
  );
}

export async function getEmployeeExperience(token: string, userId?: string): Promise<GraphResponse<unknown>> {
  try {
    return await callGraphAt<unknown>(GRAPH_BETA_URL, token, employeeExperiencePath(userId));
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get employeeExperience');
  }
}

export async function patchEmployeeExperience(
  token: string,
  body: unknown,
  userId?: string
): Promise<GraphResponse<unknown>> {
  try {
    return await callGraphAt<unknown>(GRAPH_BETA_URL, token, employeeExperiencePath(userId), {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    });
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to patch employeeExperience');
  }
}

export async function deleteEmployeeExperience(
  token: string,
  userId?: string,
  ifMatch?: string
): Promise<GraphResponse<void>> {
  const headers: Record<string, string> = {};
  if (ifMatch?.trim()) headers['If-Match'] = ifMatch.trim();
  try {
    return await callGraphAt<void>(
      GRAPH_BETA_URL,
      token,
      employeeExperiencePath(userId),
      { method: 'DELETE', headers },
      false
    );
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to delete employeeExperience');
  }
}

export async function createEmployeeExperienceAssignedRole(
  token: string,
  body: unknown,
  userId?: string
): Promise<GraphResponse<unknown>> {
  try {
    return await callGraphAt<unknown>(GRAPH_BETA_URL, token, assignedRolesPath(userId), {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    });
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to create assignedRole');
  }
}

export async function getEmployeeExperienceAssignedRole(
  token: string,
  roleId: string,
  userId?: string
): Promise<GraphResponse<unknown>> {
  try {
    return await callGraphAt<unknown>(GRAPH_BETA_URL, token, assignedRoleItemPath(userId, roleId));
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get assignedRole');
  }
}

export async function patchEmployeeExperienceAssignedRole(
  token: string,
  roleId: string,
  body: unknown,
  userId?: string
): Promise<GraphResponse<unknown>> {
  try {
    return await callGraphAt<unknown>(GRAPH_BETA_URL, token, assignedRoleItemPath(userId, roleId), {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    });
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to patch assignedRole');
  }
}

export async function deleteEmployeeExperienceAssignedRole(
  token: string,
  roleId: string,
  userId?: string,
  ifMatch?: string
): Promise<GraphResponse<void>> {
  const headers: Record<string, string> = {};
  if (ifMatch?.trim()) headers['If-Match'] = ifMatch.trim();
  try {
    return await callGraphAt<void>(
      GRAPH_BETA_URL,
      token,
      assignedRoleItemPath(userId, roleId),
      { method: 'DELETE', headers },
      false
    );
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to delete assignedRole');
  }
}

export async function listEmployeeExperienceAssignedRoleMembers(
  token: string,
  roleId: string,
  userId?: string,
  listQuery: string = ''
): Promise<GraphResponse<unknown[]>> {
  const q = listQuery.startsWith('?') ? listQuery : listQuery ? `?${listQuery}` : '';
  return fetchAllPages<unknown>(
    token,
    `${assignedRoleMembersCollectionPath(userId, roleId)}${q}`,
    'Failed to list assignedRole members',
    GRAPH_BETA_URL
  );
}

export async function createEmployeeExperienceAssignedRoleMember(
  token: string,
  roleId: string,
  body: unknown,
  userId?: string
): Promise<GraphResponse<unknown>> {
  try {
    return await callGraphAt<unknown>(GRAPH_BETA_URL, token, assignedRoleMembersCollectionPath(userId, roleId), {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    });
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to create assignedRole member');
  }
}

export async function getEmployeeExperienceAssignedRoleMember(
  token: string,
  roleId: string,
  memberId: string,
  userId?: string
): Promise<GraphResponse<unknown>> {
  try {
    return await callGraphAt<unknown>(GRAPH_BETA_URL, token, assignedRoleMemberItemPath(userId, roleId, memberId));
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get assignedRole member');
  }
}

export async function patchEmployeeExperienceAssignedRoleMember(
  token: string,
  roleId: string,
  memberId: string,
  body: unknown,
  userId?: string
): Promise<GraphResponse<unknown>> {
  try {
    return await callGraphAt<unknown>(GRAPH_BETA_URL, token, assignedRoleMemberItemPath(userId, roleId, memberId), {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    });
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to patch assignedRole member');
  }
}

export async function deleteEmployeeExperienceAssignedRoleMember(
  token: string,
  roleId: string,
  memberId: string,
  userId?: string,
  ifMatch?: string
): Promise<GraphResponse<void>> {
  const headers: Record<string, string> = {};
  if (ifMatch?.trim()) headers['If-Match'] = ifMatch.trim();
  try {
    return await callGraphAt<void>(
      GRAPH_BETA_URL,
      token,
      assignedRoleMemberItemPath(userId, roleId, memberId),
      { method: 'DELETE', headers },
      false
    );
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to delete assignedRole member');
  }
}

function assignedRoleMemberUserPath(userId: string | undefined, roleId: string, memberId: string): string {
  return `${assignedRoleMemberItemPath(userId, roleId, memberId)}/user`;
}

function assignedRoleMemberUserMailboxSettingsPath(
  userId: string | undefined,
  roleId: string,
  memberId: string
): string {
  return `${assignedRoleMemberUserPath(userId, roleId, memberId)}/mailboxSettings`;
}

function assignedRoleMemberUserServiceProvisioningErrorsPath(
  userId: string | undefined,
  roleId: string,
  memberId: string
): string {
  return `${assignedRoleMemberUserPath(userId, roleId, memberId)}/serviceProvisioningErrors`;
}

export async function getEmployeeExperienceAssignedRoleMemberUser(
  token: string,
  roleId: string,
  memberId: string,
  userId?: string
): Promise<GraphResponse<unknown>> {
  try {
    return await callGraphAt<unknown>(GRAPH_BETA_URL, token, assignedRoleMemberUserPath(userId, roleId, memberId));
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get assignedRole member user');
  }
}

export async function getEmployeeExperienceAssignedRoleMemberUserMailboxSettings(
  token: string,
  roleId: string,
  memberId: string,
  userId?: string
): Promise<GraphResponse<unknown>> {
  try {
    return await callGraphAt<unknown>(
      GRAPH_BETA_URL,
      token,
      assignedRoleMemberUserMailboxSettingsPath(userId, roleId, memberId)
    );
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get assignedRole member mailboxSettings');
  }
}

export async function patchEmployeeExperienceAssignedRoleMemberUserMailboxSettings(
  token: string,
  roleId: string,
  memberId: string,
  body: unknown,
  userId?: string
): Promise<GraphResponse<unknown>> {
  try {
    return await callGraphAt<unknown>(
      GRAPH_BETA_URL,
      token,
      assignedRoleMemberUserMailboxSettingsPath(userId, roleId, memberId),
      {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body)
      }
    );
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to patch assignedRole member mailboxSettings');
  }
}

export async function listEmployeeExperienceAssignedRoleMemberUserServiceProvisioningErrors(
  token: string,
  roleId: string,
  memberId: string,
  userId: string | undefined,
  listQuery: string = ''
): Promise<GraphResponse<unknown[]>> {
  const q = listQuery.startsWith('?') ? listQuery : listQuery ? `?${listQuery}` : '';
  return fetchAllPages<unknown>(
    token,
    `${assignedRoleMemberUserServiceProvisioningErrorsPath(userId, roleId, memberId)}${q}`,
    'Failed to list assignedRole member serviceProvisioningErrors',
    GRAPH_BETA_URL
  );
}

export async function listLearningCourseActivities(
  token: string,
  userId?: string,
  listQuery: string = ''
): Promise<GraphResponse<unknown[]>> {
  const q = listQuery.startsWith('?') ? listQuery : listQuery ? `?${listQuery}` : '';
  return fetchAllPages<unknown>(
    token,
    `${learningCourseActivitiesCollectionPath(userId)}${q}`,
    'Failed to list learningCourseActivities',
    GRAPH_BETA_URL
  );
}

export async function getLearningCourseActivity(
  token: string,
  activityId: string,
  userId?: string
): Promise<GraphResponse<unknown>> {
  const path = `${learningCourseActivitiesCollectionPath(userId)}/${encodeURIComponent(activityId.trim())}`;
  try {
    return await callGraphAt<unknown>(GRAPH_BETA_URL, token, path);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get learningCourseActivity');
  }
}

export async function getLearningCourseActivityByExternalId(
  token: string,
  externalCourseActivityId: string,
  userId?: string
): Promise<GraphResponse<unknown>> {
  const esc = escapeODataSingleQuotedKey(externalCourseActivityId.trim());
  const path = `${learningCourseActivitiesCollectionPath(userId)}(externalcourseActivityId='${esc}')`;
  try {
    return await callGraphAt<unknown>(GRAPH_BETA_URL, token, path);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get learningCourseActivity by external id');
  }
}

// --- Admin / org item insights (tenant policy) ---

export async function getAdminPeopleItemInsights(token: string): Promise<GraphResponse<unknown>> {
  try {
    return await callGraphAt<unknown>(GRAPH_BETA_URL, token, '/admin/people/itemInsights');
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get admin people itemInsights');
  }
}

export async function patchAdminPeopleItemInsights(token: string, body: unknown): Promise<GraphResponse<unknown>> {
  try {
    return await callGraphAt<unknown>(GRAPH_BETA_URL, token, '/admin/people/itemInsights', {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    });
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to patch admin people itemInsights');
  }
}

export async function deleteAdminPeopleItemInsights(token: string, ifMatch?: string): Promise<GraphResponse<void>> {
  const headers: Record<string, string> = {};
  if (ifMatch?.trim()) headers['If-Match'] = ifMatch.trim();
  try {
    return await callGraphAt<void>(
      GRAPH_BETA_URL,
      token,
      '/admin/people/itemInsights',
      { method: 'DELETE', headers },
      false
    );
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to delete admin people itemInsights');
  }
}

function organizationItemInsightsPath(organizationId: string): string {
  return `/organization/${encodeURIComponent(organizationId.trim())}/settings/itemInsights`;
}

export async function getOrganizationItemInsights(
  token: string,
  organizationId: string
): Promise<GraphResponse<unknown>> {
  try {
    return await callGraphAt<unknown>(GRAPH_BETA_URL, token, organizationItemInsightsPath(organizationId));
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get organization itemInsights');
  }
}

export async function patchOrganizationItemInsights(
  token: string,
  organizationId: string,
  body: unknown
): Promise<GraphResponse<unknown>> {
  try {
    return await callGraphAt<unknown>(GRAPH_BETA_URL, token, organizationItemInsightsPath(organizationId), {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    });
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to patch organization itemInsights');
  }
}

export async function deleteOrganizationItemInsights(
  token: string,
  organizationId: string,
  ifMatch?: string
): Promise<GraphResponse<void>> {
  const headers: Record<string, string> = {};
  if (ifMatch?.trim()) headers['If-Match'] = ifMatch.trim();
  try {
    return await callGraphAt<void>(
      GRAPH_BETA_URL,
      token,
      organizationItemInsightsPath(organizationId),
      { method: 'DELETE', headers },
      false
    );
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to delete organization itemInsights');
  }
}

// --- Work hours & locations (Viva / hybrid workplace settings) ---

function workHoursAndLocationsPath(userId?: string): string {
  const u = userId?.trim();
  if (u) return `/users/${encodeURIComponent(u)}/settings/workHoursAndLocations`;
  return '/me/settings/workHoursAndLocations';
}

export async function getWorkHoursAndLocations(token: string, userId?: string): Promise<GraphResponse<unknown>> {
  try {
    return await callGraphAt<unknown>(GRAPH_BETA_URL, token, workHoursAndLocationsPath(userId));
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get workHoursAndLocations');
  }
}

export async function patchWorkHoursAndLocations(
  token: string,
  body: unknown,
  userId?: string
): Promise<GraphResponse<unknown>> {
  try {
    return await callGraphAt<unknown>(GRAPH_BETA_URL, token, workHoursAndLocationsPath(userId), {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    });
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to patch workHoursAndLocations');
  }
}

export async function listWorkHoursOccurrences(
  token: string,
  userId: string | undefined,
  listQuery: string
): Promise<GraphResponse<unknown[]>> {
  const q = listQuery.startsWith('?') ? listQuery : listQuery ? `?${listQuery}` : '';
  return fetchAllPages<unknown>(
    token,
    `${workHoursAndLocationsPath(userId)}/occurrences${q}`,
    'Failed to list workHours occurrences',
    GRAPH_BETA_URL
  );
}

export async function postWorkHoursSetCurrentLocation(
  token: string,
  userId?: string,
  body?: unknown
): Promise<GraphResponse<unknown>> {
  const path = `${workHoursAndLocationsPath(userId)}/occurrences/setCurrentLocation`;
  try {
    return await callGraphAt<unknown>(GRAPH_BETA_URL, token, path, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: body !== undefined ? JSON.stringify(body) : '{}'
    });
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to setCurrentLocation');
  }
}

export async function getWorkHoursOccurrence(
  token: string,
  occurrenceId: string,
  userId?: string
): Promise<GraphResponse<unknown>> {
  const path = `${workHoursAndLocationsPath(userId)}/occurrences/${encodeURIComponent(occurrenceId.trim())}`;
  try {
    return await callGraphAt<unknown>(GRAPH_BETA_URL, token, path);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get workPlanOccurrence');
  }
}

export async function patchWorkHoursOccurrence(
  token: string,
  occurrenceId: string,
  body: unknown,
  userId?: string
): Promise<GraphResponse<unknown>> {
  const path = `${workHoursAndLocationsPath(userId)}/occurrences/${encodeURIComponent(occurrenceId.trim())}`;
  try {
    return await callGraphAt<unknown>(GRAPH_BETA_URL, token, path, {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    });
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to patch workPlanOccurrence');
  }
}

export async function deleteWorkHoursOccurrence(
  token: string,
  occurrenceId: string,
  userId?: string,
  ifMatch?: string
): Promise<GraphResponse<void>> {
  const headers: Record<string, string> = {};
  if (ifMatch?.trim()) headers['If-Match'] = ifMatch.trim();
  const path = `${workHoursAndLocationsPath(userId)}/occurrences/${encodeURIComponent(occurrenceId.trim())}`;
  try {
    return await callGraphAt<void>(GRAPH_BETA_URL, token, path, { method: 'DELETE', headers }, false);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to delete workPlanOccurrence');
  }
}

export async function getWorkHoursOccurrencesView(
  token: string,
  startDateTime: string,
  endDateTime: string,
  userId?: string
): Promise<GraphResponse<unknown>> {
  const s = escapeODataSingleQuotedKey(startDateTime.trim());
  const e = escapeODataSingleQuotedKey(endDateTime.trim());
  const path = `${workHoursAndLocationsPath(userId)}/occurrencesView(startDateTime='${s}',endDateTime='${e}')`;
  try {
    return await callGraphAt<unknown>(GRAPH_BETA_URL, token, path);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get occurrencesView');
  }
}

export async function listWorkHoursRecurrences(
  token: string,
  userId: string | undefined,
  listQuery: string
): Promise<GraphResponse<unknown[]>> {
  const q = listQuery.startsWith('?') ? listQuery : listQuery ? `?${listQuery}` : '';
  return fetchAllPages<unknown>(
    token,
    `${workHoursAndLocationsPath(userId)}/recurrences${q}`,
    'Failed to list workHours recurrences',
    GRAPH_BETA_URL
  );
}

export async function getWorkHoursRecurrence(
  token: string,
  recurrenceId: string,
  userId?: string
): Promise<GraphResponse<unknown>> {
  const path = `${workHoursAndLocationsPath(userId)}/recurrences/${encodeURIComponent(recurrenceId.trim())}`;
  try {
    return await callGraphAt<unknown>(GRAPH_BETA_URL, token, path);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get workPlanRecurrence');
  }
}

export async function patchWorkHoursRecurrence(
  token: string,
  recurrenceId: string,
  body: unknown,
  userId?: string
): Promise<GraphResponse<unknown>> {
  const path = `${workHoursAndLocationsPath(userId)}/recurrences/${encodeURIComponent(recurrenceId.trim())}`;
  try {
    return await callGraphAt<unknown>(GRAPH_BETA_URL, token, path, {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    });
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to patch workPlanRecurrence');
  }
}

export async function deleteWorkHoursRecurrence(
  token: string,
  recurrenceId: string,
  userId?: string,
  ifMatch?: string
): Promise<GraphResponse<void>> {
  const headers: Record<string, string> = {};
  if (ifMatch?.trim()) headers['If-Match'] = ifMatch.trim();
  const path = `${workHoursAndLocationsPath(userId)}/recurrences/${encodeURIComponent(recurrenceId.trim())}`;
  try {
    return await callGraphAt<void>(GRAPH_BETA_URL, token, path, { method: 'DELETE', headers }, false);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to delete workPlanRecurrence');
  }
}

// --- Viva Engage Q&A / online meeting conversations ---

const MEETING_CONV = '/communications/onlineMeetingConversations';

export async function listOnlineMeetingEngagementConversations(
  token: string,
  listQuery: string = ''
): Promise<GraphResponse<unknown[]>> {
  const q = listQuery.startsWith('?') ? listQuery : listQuery ? `?${listQuery}` : '';
  return fetchAllPages<unknown>(
    token,
    `${MEETING_CONV}${q}`,
    'Failed to list onlineMeetingConversations',
    GRAPH_BETA_URL
  );
}

export async function createOnlineMeetingEngagementConversation(
  token: string,
  body: unknown
): Promise<GraphResponse<unknown>> {
  try {
    return await callGraphAt<unknown>(GRAPH_BETA_URL, token, MEETING_CONV, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    });
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to create onlineMeetingConversation');
  }
}

export async function getOnlineMeetingEngagementConversation(
  token: string,
  conversationId: string
): Promise<GraphResponse<unknown>> {
  const path = `${MEETING_CONV}/${encodeURIComponent(conversationId.trim())}`;
  try {
    return await callGraphAt<unknown>(GRAPH_BETA_URL, token, path);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get onlineMeetingConversation');
  }
}

export async function patchOnlineMeetingEngagementConversation(
  token: string,
  conversationId: string,
  body: unknown
): Promise<GraphResponse<unknown>> {
  const path = `${MEETING_CONV}/${encodeURIComponent(conversationId.trim())}`;
  try {
    return await callGraphAt<unknown>(GRAPH_BETA_URL, token, path, {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    });
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to patch onlineMeetingConversation');
  }
}

export async function deleteOnlineMeetingEngagementConversation(
  token: string,
  conversationId: string,
  ifMatch?: string
): Promise<GraphResponse<void>> {
  const headers: Record<string, string> = {};
  if (ifMatch?.trim()) headers['If-Match'] = ifMatch.trim();
  const path = `${MEETING_CONV}/${encodeURIComponent(conversationId.trim())}`;
  try {
    return await callGraphAt<void>(GRAPH_BETA_URL, token, path, { method: 'DELETE', headers }, false);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to delete onlineMeetingConversation');
  }
}

export async function listOnlineMeetingEngagementConversationMessages(
  token: string,
  conversationId: string,
  listQuery: string = ''
): Promise<GraphResponse<unknown[]>> {
  const q = listQuery.startsWith('?') ? listQuery : listQuery ? `?${listQuery}` : '';
  return fetchAllPages<unknown>(
    token,
    `${MEETING_CONV}/${encodeURIComponent(conversationId.trim())}/messages${q}`,
    'Failed to list engagement conversation messages',
    GRAPH_BETA_URL
  );
}

export async function getOnlineMeetingEngagementConversationMessage(
  token: string,
  conversationId: string,
  messageId: string
): Promise<GraphResponse<unknown>> {
  const path = `${MEETING_CONV}/${encodeURIComponent(conversationId.trim())}/messages/${encodeURIComponent(messageId.trim())}`;
  try {
    return await callGraphAt<unknown>(GRAPH_BETA_URL, token, path);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get engagement conversation message');
  }
}
