/**
 * Microsoft Graph **beta** tenant **`/employeeExperience`** APIs (Viva org-wide / admin scenarios).
 * Delegated vs application permissions vary by operation — see docs/GRAPH_SCOPES.md.
 */

import {
  callGraphAt,
  callGraphAtText,
  fetchAllPages,
  GraphApiError,
  type GraphResponse,
  graphError
} from './graph-client.js';
import { getGraphBetaUrl } from './graph-constants.js';
import { escapeODataSingleQuotedKey } from './graph-viva-client.js';

const EX = '/employeeExperience';

function listSuffix(listQuery: string): string {
  if (!listQuery) return '';
  return listQuery.startsWith('?') ? listQuery : `?${listQuery}`;
}

async function betaGet(token: string, path: string): Promise<GraphResponse<unknown>> {
  try {
    return await callGraphAt<unknown>(getGraphBetaUrl(), token, path);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Graph GET failed');
  }
}

async function betaPatch(token: string, path: string, body: unknown): Promise<GraphResponse<unknown>> {
  try {
    return await callGraphAt<unknown>(getGraphBetaUrl(), token, path, {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    });
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Graph PATCH failed');
  }
}

async function betaPost(token: string, path: string, body: unknown): Promise<GraphResponse<unknown>> {
  try {
    return await callGraphAt<unknown>(getGraphBetaUrl(), token, path, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    });
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Graph POST failed');
  }
}

async function betaDelete(token: string, path: string, ifMatch?: string): Promise<GraphResponse<void>> {
  const headers: Record<string, string> = {};
  if (ifMatch?.trim()) headers['If-Match'] = ifMatch.trim();
  try {
    return await callGraphAt<void>(getGraphBetaUrl(), token, path, { method: 'DELETE', headers }, false);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Graph DELETE failed');
  }
}

async function betaList(
  token: string,
  path: string,
  listQuery: string,
  errMsg: string
): Promise<GraphResponse<unknown[]>> {
  return fetchAllPages<unknown>(token, `${path}${listSuffix(listQuery)}`, errMsg, getGraphBetaUrl());
}

// --- /employeeExperience singleton ---

export async function getTenantEmployeeExperience(token: string): Promise<GraphResponse<unknown>> {
  return betaGet(token, EX);
}

export async function patchTenantEmployeeExperience(token: string, body: unknown): Promise<GraphResponse<unknown>> {
  return betaPatch(token, EX, body);
}

export async function deleteTenantEmployeeExperience(token: string, ifMatch?: string): Promise<GraphResponse<void>> {
  return betaDelete(token, EX, ifMatch);
}

// --- communities ---

export async function listTenantCommunities(token: string, listQuery: string = ''): Promise<GraphResponse<unknown[]>> {
  return betaList(token, `${EX}/communities`, listQuery, 'Failed to list communities');
}

export async function createTenantCommunity(token: string, body: unknown): Promise<GraphResponse<unknown>> {
  return betaPost(token, `${EX}/communities`, body);
}

export async function getTenantCommunity(token: string, communityId: string): Promise<GraphResponse<unknown>> {
  return betaGet(token, `${EX}/communities/${encodeURIComponent(communityId.trim())}`);
}

export async function patchTenantCommunity(
  token: string,
  communityId: string,
  body: unknown
): Promise<GraphResponse<unknown>> {
  return betaPatch(token, `${EX}/communities/${encodeURIComponent(communityId.trim())}`, body);
}

export async function deleteTenantCommunity(
  token: string,
  communityId: string,
  ifMatch?: string
): Promise<GraphResponse<void>> {
  return betaDelete(token, `${EX}/communities/${encodeURIComponent(communityId.trim())}`, ifMatch);
}

export async function getTenantCommunityGroup(token: string, communityId: string): Promise<GraphResponse<unknown>> {
  return betaGet(token, `${EX}/communities/${encodeURIComponent(communityId.trim())}/group`);
}

export async function listTenantCommunityOwners(
  token: string,
  communityId: string,
  listQuery: string = ''
): Promise<GraphResponse<unknown[]>> {
  return betaList(
    token,
    `${EX}/communities/${encodeURIComponent(communityId.trim())}/owners`,
    listQuery,
    'Failed to list community owners'
  );
}

export async function getTenantCommunityOwner(
  token: string,
  communityId: string,
  ownerUserId: string
): Promise<GraphResponse<unknown>> {
  return betaGet(
    token,
    `${EX}/communities/${encodeURIComponent(communityId.trim())}/owners/${encodeURIComponent(ownerUserId.trim())}`
  );
}

function tenantCommunityOwnersCollection(communityId: string): string {
  return `${EX}/communities/${encodeURIComponent(communityId.trim())}/owners`;
}

function tenantCommunityOwnerByIdPath(communityId: string, ownerUserId: string): string {
  return `${tenantCommunityOwnersCollection(communityId)}/${encodeURIComponent(ownerUserId.trim())}`;
}

export async function getTenantCommunityOwnerByUserPrincipalName(
  token: string,
  communityId: string,
  userPrincipalName: string
): Promise<GraphResponse<unknown>> {
  const esc = escapeODataSingleQuotedKey(userPrincipalName.trim());
  return betaGet(token, `${tenantCommunityOwnersCollection(communityId)}(userPrincipalName='${esc}')`);
}

export async function getTenantCommunityOwnerMailboxSettings(
  token: string,
  communityId: string,
  ownerUserId: string
): Promise<GraphResponse<unknown>> {
  return betaGet(token, `${tenantCommunityOwnerByIdPath(communityId, ownerUserId)}/mailboxSettings`);
}

export async function patchTenantCommunityOwnerMailboxSettings(
  token: string,
  communityId: string,
  ownerUserId: string,
  body: unknown
): Promise<GraphResponse<unknown>> {
  return betaPatch(token, `${tenantCommunityOwnerByIdPath(communityId, ownerUserId)}/mailboxSettings`, body);
}

export async function listTenantCommunityOwnerServiceProvisioningErrors(
  token: string,
  communityId: string,
  ownerUserId: string,
  listQuery: string = ''
): Promise<GraphResponse<unknown[]>> {
  return betaList(
    token,
    `${tenantCommunityOwnerByIdPath(communityId, ownerUserId)}/serviceProvisioningErrors`,
    listQuery,
    'Failed to list community owner serviceProvisioningErrors'
  );
}

// --- engagement async operations ---

export async function listTenantEngagementAsyncOperations(
  token: string,
  listQuery: string = ''
): Promise<GraphResponse<unknown[]>> {
  return betaList(token, `${EX}/engagementAsyncOperations`, listQuery, 'Failed to list engagementAsyncOperations');
}

export async function createTenantEngagementAsyncOperation(
  token: string,
  body: unknown
): Promise<GraphResponse<unknown>> {
  return betaPost(token, `${EX}/engagementAsyncOperations`, body);
}

export async function getTenantEngagementAsyncOperation(
  token: string,
  operationId: string
): Promise<GraphResponse<unknown>> {
  return betaGet(token, `${EX}/engagementAsyncOperations/${encodeURIComponent(operationId.trim())}`);
}

export async function patchTenantEngagementAsyncOperation(
  token: string,
  operationId: string,
  body: unknown
): Promise<GraphResponse<unknown>> {
  return betaPatch(token, `${EX}/engagementAsyncOperations/${encodeURIComponent(operationId.trim())}`, body);
}

export async function deleteTenantEngagementAsyncOperation(
  token: string,
  operationId: string,
  ifMatch?: string
): Promise<GraphResponse<void>> {
  return betaDelete(token, `${EX}/engagementAsyncOperations/${encodeURIComponent(operationId.trim())}`, ifMatch);
}

// --- goals ---

export async function getTenantGoals(token: string): Promise<GraphResponse<unknown>> {
  return betaGet(token, `${EX}/goals`);
}

export async function patchTenantGoals(token: string, body: unknown): Promise<GraphResponse<unknown>> {
  return betaPatch(token, `${EX}/goals`, body);
}

export async function deleteTenantGoals(token: string, ifMatch?: string): Promise<GraphResponse<void>> {
  return betaDelete(token, `${EX}/goals`, ifMatch);
}

export async function listTenantGoalsExportJobs(
  token: string,
  listQuery: string = ''
): Promise<GraphResponse<unknown[]>> {
  return betaList(token, `${EX}/goals/exportJobs`, listQuery, 'Failed to list goals exportJobs');
}

export async function createTenantGoalsExportJob(token: string, body: unknown): Promise<GraphResponse<unknown>> {
  return betaPost(token, `${EX}/goals/exportJobs`, body);
}

export async function getTenantGoalsExportJob(token: string, jobId: string): Promise<GraphResponse<unknown>> {
  return betaGet(token, `${EX}/goals/exportJobs/${encodeURIComponent(jobId.trim())}`);
}

export async function patchTenantGoalsExportJob(
  token: string,
  jobId: string,
  body: unknown
): Promise<GraphResponse<unknown>> {
  return betaPatch(token, `${EX}/goals/exportJobs/${encodeURIComponent(jobId.trim())}`, body);
}

export async function deleteTenantGoalsExportJob(
  token: string,
  jobId: string,
  ifMatch?: string
): Promise<GraphResponse<void>> {
  return betaDelete(token, `${EX}/goals/exportJobs/${encodeURIComponent(jobId.trim())}`, ifMatch);
}

export async function getTenantGoalsExportJobContent(token: string, jobId: string): Promise<GraphResponse<string>> {
  try {
    return await callGraphAtText(
      getGraphBetaUrl(),
      token,
      `${EX}/goals/exportJobs/${encodeURIComponent(jobId.trim())}/content`
    );
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get export job content');
  }
}

// --- tenant learning course activities ---

export async function listTenantRootLearningCourseActivities(
  token: string,
  listQuery: string = ''
): Promise<GraphResponse<unknown[]>> {
  return betaList(token, `${EX}/learningCourseActivities`, listQuery, 'Failed to list tenant learningCourseActivities');
}

export async function createTenantRootLearningCourseActivity(
  token: string,
  body: unknown
): Promise<GraphResponse<unknown>> {
  return betaPost(token, `${EX}/learningCourseActivities`, body);
}

export async function getTenantRootLearningCourseActivity(
  token: string,
  activityId: string
): Promise<GraphResponse<unknown>> {
  return betaGet(token, `${EX}/learningCourseActivities/${encodeURIComponent(activityId.trim())}`);
}

export async function getTenantRootLearningCourseActivityByExternal(
  token: string,
  externalCourseActivityId: string
): Promise<GraphResponse<unknown>> {
  const esc = escapeODataSingleQuotedKey(externalCourseActivityId.trim());
  return betaGet(token, `${EX}/learningCourseActivities(externalcourseActivityId='${esc}')`);
}

export async function patchTenantRootLearningCourseActivity(
  token: string,
  activityId: string,
  body: unknown
): Promise<GraphResponse<unknown>> {
  return betaPatch(token, `${EX}/learningCourseActivities/${encodeURIComponent(activityId.trim())}`, body);
}

export async function deleteTenantRootLearningCourseActivity(
  token: string,
  activityId: string,
  ifMatch?: string
): Promise<GraphResponse<void>> {
  return betaDelete(token, `${EX}/learningCourseActivities/${encodeURIComponent(activityId.trim())}`, ifMatch);
}

// --- learning providers ---

export async function listTenantLearningProviders(
  token: string,
  listQuery: string = ''
): Promise<GraphResponse<unknown[]>> {
  return betaList(token, `${EX}/learningProviders`, listQuery, 'Failed to list learningProviders');
}

export async function createTenantLearningProvider(token: string, body: unknown): Promise<GraphResponse<unknown>> {
  return betaPost(token, `${EX}/learningProviders`, body);
}

export async function getTenantLearningProvider(token: string, providerId: string): Promise<GraphResponse<unknown>> {
  return betaGet(token, `${EX}/learningProviders/${encodeURIComponent(providerId.trim())}`);
}

export async function patchTenantLearningProvider(
  token: string,
  providerId: string,
  body: unknown
): Promise<GraphResponse<unknown>> {
  return betaPatch(token, `${EX}/learningProviders/${encodeURIComponent(providerId.trim())}`, body);
}

export async function deleteTenantLearningProvider(
  token: string,
  providerId: string,
  ifMatch?: string
): Promise<GraphResponse<void>> {
  return betaDelete(token, `${EX}/learningProviders/${encodeURIComponent(providerId.trim())}`, ifMatch);
}

function learningContentsBase(providerId: string): string {
  return `${EX}/learningProviders/${encodeURIComponent(providerId.trim())}/learningContents`;
}

export async function listTenantLearningContents(
  token: string,
  providerId: string,
  listQuery: string = ''
): Promise<GraphResponse<unknown[]>> {
  return betaList(token, learningContentsBase(providerId), listQuery, 'Failed to list learningContents');
}

export async function createTenantLearningContent(
  token: string,
  providerId: string,
  body: unknown
): Promise<GraphResponse<unknown>> {
  return betaPost(token, learningContentsBase(providerId), body);
}

export async function getTenantLearningContent(
  token: string,
  providerId: string,
  contentId: string
): Promise<GraphResponse<unknown>> {
  return betaGet(token, `${learningContentsBase(providerId)}/${encodeURIComponent(contentId.trim())}`);
}

export async function getTenantLearningContentByExternal(
  token: string,
  providerId: string,
  externalId: string
): Promise<GraphResponse<unknown>> {
  const esc = escapeODataSingleQuotedKey(externalId.trim());
  return betaGet(token, `${learningContentsBase(providerId)}(externalId='${esc}')`);
}

export async function patchTenantLearningContent(
  token: string,
  providerId: string,
  contentId: string,
  body: unknown
): Promise<GraphResponse<unknown>> {
  return betaPatch(token, `${learningContentsBase(providerId)}/${encodeURIComponent(contentId.trim())}`, body);
}

export async function deleteTenantLearningContent(
  token: string,
  providerId: string,
  contentId: string,
  ifMatch?: string
): Promise<GraphResponse<void>> {
  return betaDelete(token, `${learningContentsBase(providerId)}/${encodeURIComponent(contentId.trim())}`, ifMatch);
}

function providerActivitiesBase(providerId: string): string {
  return `${EX}/learningProviders/${encodeURIComponent(providerId.trim())}/learningCourseActivities`;
}

export async function listTenantProviderLearningCourseActivities(
  token: string,
  providerId: string,
  listQuery: string = ''
): Promise<GraphResponse<unknown[]>> {
  return betaList(
    token,
    providerActivitiesBase(providerId),
    listQuery,
    'Failed to list provider learningCourseActivities'
  );
}

export async function createTenantProviderLearningCourseActivity(
  token: string,
  providerId: string,
  body: unknown
): Promise<GraphResponse<unknown>> {
  return betaPost(token, providerActivitiesBase(providerId), body);
}

export async function getTenantProviderLearningCourseActivity(
  token: string,
  providerId: string,
  activityId: string
): Promise<GraphResponse<unknown>> {
  return betaGet(token, `${providerActivitiesBase(providerId)}/${encodeURIComponent(activityId.trim())}`);
}

export async function getTenantProviderLearningCourseActivityByExternal(
  token: string,
  providerId: string,
  externalCourseActivityId: string
): Promise<GraphResponse<unknown>> {
  const esc = escapeODataSingleQuotedKey(externalCourseActivityId.trim());
  return betaGet(token, `${providerActivitiesBase(providerId)}(externalcourseActivityId='${esc}')`);
}

export async function patchTenantProviderLearningCourseActivity(
  token: string,
  providerId: string,
  activityId: string,
  body: unknown
): Promise<GraphResponse<unknown>> {
  return betaPatch(token, `${providerActivitiesBase(providerId)}/${encodeURIComponent(activityId.trim())}`, body);
}

export async function deleteTenantProviderLearningCourseActivity(
  token: string,
  providerId: string,
  activityId: string,
  ifMatch?: string
): Promise<GraphResponse<void>> {
  return betaDelete(token, `${providerActivitiesBase(providerId)}/${encodeURIComponent(activityId.trim())}`, ifMatch);
}

// --- tenant engagement roles (catalog) ---

export async function listTenantEngagementRoles(
  token: string,
  listQuery: string = ''
): Promise<GraphResponse<unknown[]>> {
  return betaList(token, `${EX}/roles`, listQuery, 'Failed to list tenant engagement roles');
}

export async function createTenantEngagementRole(token: string, body: unknown): Promise<GraphResponse<unknown>> {
  return betaPost(token, `${EX}/roles`, body);
}

export async function getTenantEngagementRole(token: string, roleId: string): Promise<GraphResponse<unknown>> {
  return betaGet(token, `${EX}/roles/${encodeURIComponent(roleId.trim())}`);
}

export async function patchTenantEngagementRole(
  token: string,
  roleId: string,
  body: unknown
): Promise<GraphResponse<unknown>> {
  return betaPatch(token, `${EX}/roles/${encodeURIComponent(roleId.trim())}`, body);
}

export async function deleteTenantEngagementRole(
  token: string,
  roleId: string,
  ifMatch?: string
): Promise<GraphResponse<void>> {
  return betaDelete(token, `${EX}/roles/${encodeURIComponent(roleId.trim())}`, ifMatch);
}

export async function listTenantEngagementRoleMembers(
  token: string,
  roleId: string,
  listQuery: string = ''
): Promise<GraphResponse<unknown[]>> {
  return betaList(
    token,
    `${EX}/roles/${encodeURIComponent(roleId.trim())}/members`,
    listQuery,
    'Failed to list tenant role members'
  );
}

export async function createTenantEngagementRoleMember(
  token: string,
  roleId: string,
  body: unknown
): Promise<GraphResponse<unknown>> {
  return betaPost(token, `${EX}/roles/${encodeURIComponent(roleId.trim())}/members`, body);
}

export async function getTenantEngagementRoleMember(
  token: string,
  roleId: string,
  memberId: string
): Promise<GraphResponse<unknown>> {
  return betaGet(
    token,
    `${EX}/roles/${encodeURIComponent(roleId.trim())}/members/${encodeURIComponent(memberId.trim())}`
  );
}

export async function patchTenantEngagementRoleMember(
  token: string,
  roleId: string,
  memberId: string,
  body: unknown
): Promise<GraphResponse<unknown>> {
  return betaPatch(
    token,
    `${EX}/roles/${encodeURIComponent(roleId.trim())}/members/${encodeURIComponent(memberId.trim())}`,
    body
  );
}

export async function deleteTenantEngagementRoleMember(
  token: string,
  roleId: string,
  memberId: string,
  ifMatch?: string
): Promise<GraphResponse<void>> {
  return betaDelete(
    token,
    `${EX}/roles/${encodeURIComponent(roleId.trim())}/members/${encodeURIComponent(memberId.trim())}`,
    ifMatch
  );
}

function tenantRoleMemberPath(roleId: string, memberId: string): string {
  return `${EX}/roles/${encodeURIComponent(roleId.trim())}/members/${encodeURIComponent(memberId.trim())}`;
}

export async function getTenantEngagementRoleMemberUser(
  token: string,
  roleId: string,
  memberId: string
): Promise<GraphResponse<unknown>> {
  return betaGet(token, `${tenantRoleMemberPath(roleId, memberId)}/user`);
}

export async function getTenantEngagementRoleMemberUserMailboxSettings(
  token: string,
  roleId: string,
  memberId: string
): Promise<GraphResponse<unknown>> {
  return betaGet(token, `${tenantRoleMemberPath(roleId, memberId)}/user/mailboxSettings`);
}

export async function patchTenantEngagementRoleMemberUserMailboxSettings(
  token: string,
  roleId: string,
  memberId: string,
  body: unknown
): Promise<GraphResponse<unknown>> {
  return betaPatch(token, `${tenantRoleMemberPath(roleId, memberId)}/user/mailboxSettings`, body);
}

export async function listTenantEngagementRoleMemberUserServiceProvisioningErrors(
  token: string,
  roleId: string,
  memberId: string,
  listQuery: string = ''
): Promise<GraphResponse<unknown[]>> {
  return betaList(
    token,
    `${tenantRoleMemberPath(roleId, memberId)}/user/serviceProvisioningErrors`,
    listQuery,
    'Failed to list tenant role member serviceProvisioningErrors'
  );
}
