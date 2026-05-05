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
  userId?: string
): Promise<GraphResponse<unknown[]>> {
  return fetchAllPages<unknown>(
    token,
    assignedRolesPath(userId),
    'Failed to list employeeExperience assignedRoles',
    GRAPH_BETA_URL
  );
}
