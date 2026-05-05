/**
 * Microsoft 365 Copilot APIs on Microsoft Graph (`/copilot/...`).
 * @see https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/copilot-apis-overview
 */

import { graphInvoke, graphInvokeText } from './graph-advanced-client.js';
import { callGraphAbsolute, GraphApiError, type GraphResponse, graphError, graphResult } from './graph-client.js';
import { GRAPH_BETA_URL } from './graph-constants.js';

/** OData-style Copilot chat actions (canonical Graph paths). */
export const COPILOT_CONVERSATION_CHAT_PATH_SUFFIX = '/microsoft.graph.copilot.chat';
export const COPILOT_CONVERSATION_CHAT_STREAM_PATH_SUFFIX = '/microsoft.graph.copilot.chatOverStream';

const REALTIME_ACTIVITY = '/copilot/communications/realtimeActivityFeed';
const COPILOT_COMMUNICATIONS = '/copilot/communications';

function aiUserOnlineMeetingBase(userId: string, onlineMeetingId: string): string {
  const u = copilotUserSegment(userId);
  const m = encodeURIComponent(onlineMeetingId.trim());
  return `/copilot/users/${u}/onlineMeetings/${m}`;
}

export const COPILOT_REPORT_PERIODS = ['D7', 'D30', 'D90', 'D180', 'ALL'] as const;
export type CopilotReportPeriod = (typeof COPILOT_REPORT_PERIODS)[number];

const MAX_SEARCH_QUERY = 1500;
const MAX_RETRIEVAL_QUERY = 1500;

export const COPILOT_RETRIEVAL_DATA_SOURCES = ['sharePoint', 'oneDriveBusiness', 'externalItem'] as const;
export type CopilotRetrievalDataSource = (typeof COPILOT_RETRIEVAL_DATA_SOURCES)[number];

function assertCopilotRetrievalDataSource(s: string): CopilotRetrievalDataSource {
  const t = s.trim();
  if (!(COPILOT_RETRIEVAL_DATA_SOURCES as readonly string[]).includes(t)) {
    throw new GraphApiError(
      `dataSource must be one of: ${COPILOT_RETRIEVAL_DATA_SOURCES.join(', ')}`,
      'InvalidDataSource',
      400
    );
  }
  return t as CopilotRetrievalDataSource;
}

export function buildCopilotRetrievalBody(opts: {
  queryString: string;
  dataSource: string;
  filterExpression?: string;
  maximumNumberOfResults?: number;
  resourceMetadata?: string[];
}): Record<string, unknown> {
  const q = opts.queryString.trim();
  if (!q) throw new GraphApiError('queryString is required', 'InvalidQuery', 400);
  if (q.length > MAX_RETRIEVAL_QUERY) {
    throw new GraphApiError(`queryString exceeds ${MAX_RETRIEVAL_QUERY} characters`, 'InvalidQuery', 400);
  }
  const body: Record<string, unknown> = {
    queryString: q,
    dataSource: assertCopilotRetrievalDataSource(opts.dataSource)
  };
  if (opts.filterExpression?.trim()) body.filterExpression = opts.filterExpression.trim();
  if (opts.maximumNumberOfResults !== undefined) {
    const n = opts.maximumNumberOfResults;
    if (Number.isNaN(n) || n < 1 || n > 25) {
      throw new GraphApiError('maximumNumberOfResults must be 1–25', 'InvalidMax', 400);
    }
    body.maximumNumberOfResults = n;
  }
  if (opts.resourceMetadata && opts.resourceMetadata.length > 0) {
    body.resourceMetadata = opts.resourceMetadata;
  }
  return body;
}

export async function copilotRetrieval(
  token: string,
  body: Record<string, unknown>,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  return graphInvoke(token, { method: 'POST', path: '/copilot/retrieval', body, beta });
}

export function assertCopilotReportPeriod(period: string): CopilotReportPeriod {
  const p = period.trim() as CopilotReportPeriod;
  if (!(COPILOT_REPORT_PERIODS as readonly string[]).includes(p)) {
    throw new GraphApiError(`period must be one of: ${COPILOT_REPORT_PERIODS.join(', ')}`, 'InvalidPeriod', 400);
  }
  return p;
}

export type CopilotReportFunctionName =
  | 'getMicrosoft365CopilotUserCountSummary'
  | 'getMicrosoft365CopilotUserCountTrend'
  | 'getMicrosoft365CopilotUsageUserDetail';

/** OData function-style path for Copilot usage report endpoints. */
export function copilotReportPath(fn: CopilotReportFunctionName, period: string): string {
  const p = assertCopilotReportPeriod(period);
  return `/copilot/reports/${fn}(period='${p}')?$format=application/json`;
}

export function copilotUserSegment(userId: string): string {
  const id = userId.trim();
  if (!id) throw new GraphApiError('user id is required', 'InvalidUser', 400);
  return encodeURIComponent(id);
}

/** Append OData query to a path (handles existing `?`). */
export function appendCopilotODataQuery(path: string, odataQuery: string | undefined): string {
  const q = odataQuery?.trim();
  if (!q) return path;
  const qs = q.startsWith('?') ? q.slice(1) : q;
  return path.includes('?') ? `${path}&${qs}` : `${path}?${qs}`;
}

export function buildCopilotSearchBody(opts: {
  query: string;
  pageSize?: number;
  oneDriveFilterExpression?: string;
  resourceMetadataNames?: string[];
}): Record<string, unknown> {
  const q = opts.query.trim();
  if (!q) throw new GraphApiError('query is required', 'InvalidQuery', 400);
  if (q.length > MAX_SEARCH_QUERY) {
    throw new GraphApiError(`query exceeds ${MAX_SEARCH_QUERY} characters`, 'InvalidQuery', 400);
  }
  const body: Record<string, unknown> = { query: q };
  if (opts.pageSize !== undefined) {
    if (Number.isNaN(opts.pageSize) || opts.pageSize < 1 || opts.pageSize > 100) {
      throw new GraphApiError('pageSize must be 1–100', 'InvalidPageSize', 400);
    }
    body.pageSize = opts.pageSize;
  }
  const filter = opts.oneDriveFilterExpression?.trim();
  const meta = opts.resourceMetadataNames?.filter(Boolean);
  if (filter || (meta && meta.length > 0)) {
    const oneDrive: Record<string, unknown> = {};
    if (filter) oneDrive.filterExpression = filter;
    if (meta && meta.length > 0) oneDrive.resourceMetadataNames = meta;
    body.dataSources = { oneDrive };
  }
  return body;
}

export async function copilotSearch(
  token: string,
  body: Record<string, unknown>,
  beta: boolean = true
): Promise<GraphResponse<unknown>> {
  return graphInvoke(token, { method: 'POST', path: '/copilot/search', body, beta });
}

export async function copilotSearchNextPage(token: string, nextLink: string): Promise<GraphResponse<unknown>> {
  try {
    return await callGraphAbsolute<unknown>(token, nextLink.trim(), { method: 'GET' }, true);
  } catch (err) {
    if (err instanceof GraphApiError) {
      return { ok: false, error: { message: err.message, code: err.code, status: err.status } };
    }
    return { ok: false, error: { message: err instanceof Error ? err.message : 'Request failed' } };
  }
}

export async function copilotConversationCreate(token: string, beta: boolean = true): Promise<GraphResponse<unknown>> {
  return graphInvoke(token, { method: 'POST', path: '/copilot/conversations', body: {}, beta });
}

export async function copilotConversationChat(
  token: string,
  conversationId: string,
  body: Record<string, unknown>,
  beta: boolean = true
): Promise<GraphResponse<unknown>> {
  const id = encodeURIComponent(conversationId.trim());
  return graphInvoke(token, {
    method: 'POST',
    path: `/copilot/conversations/${id}${COPILOT_CONVERSATION_CHAT_PATH_SUFFIX}`,
    body,
    beta
  });
}

export async function copilotConversationChatOverStream(
  token: string,
  conversationId: string,
  body: Record<string, unknown>,
  beta: boolean = true
): Promise<GraphResponse<string>> {
  const id = encodeURIComponent(conversationId.trim());
  return graphInvokeText(token, {
    method: 'POST',
    path: `/copilot/conversations/${id}${COPILOT_CONVERSATION_CHAT_STREAM_PATH_SUFFIX}`,
    body,
    beta
  });
}

export async function copilotInteractionsExportList(
  token: string,
  userId: string,
  rawODataQuery: string | undefined,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  const base = `/copilot/users/${copilotUserSegment(userId)}/interactionHistory/getAllEnterpriseInteractions()`;
  return graphInvoke(token, {
    method: 'GET',
    path: appendCopilotODataQuery(base, rawODataQuery),
    beta
  });
}

/** App-only tenant export: GET /copilot/interactionHistory/getAllEnterpriseInteractions() */
export async function copilotInteractionsTenantExportList(
  token: string,
  rawODataQuery: string | undefined,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  const base = '/copilot/interactionHistory/getAllEnterpriseInteractions()';
  return graphInvoke(token, {
    method: 'GET',
    path: appendCopilotODataQuery(base, rawODataQuery),
    beta
  });
}

export async function copilotMeetingInsightsList(
  token: string,
  userId: string,
  onlineMeetingId: string,
  odataQuery: string | undefined,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  const base = `${aiUserOnlineMeetingBase(userId, onlineMeetingId)}/aiInsights`;
  return graphInvoke(token, { method: 'GET', path: appendCopilotODataQuery(base, odataQuery), beta });
}

export async function copilotMeetingInsightGet(
  token: string,
  userId: string,
  onlineMeetingId: string,
  aiInsightId: string,
  odataQuery: string | undefined,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  const base = `${aiUserOnlineMeetingBase(userId, onlineMeetingId)}/aiInsights/${encodeURIComponent(aiInsightId.trim())}`;
  return graphInvoke(token, { method: 'GET', path: appendCopilotODataQuery(base, odataQuery), beta });
}

export async function copilotMeetingAiInsightsCreate(
  token: string,
  userId: string,
  onlineMeetingId: string,
  body: Record<string, unknown>,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  const path = `${aiUserOnlineMeetingBase(userId, onlineMeetingId)}/aiInsights`;
  return graphInvoke(token, { method: 'POST', path, body, beta });
}

export async function copilotMeetingAiInsightsCount(
  token: string,
  userId: string,
  onlineMeetingId: string,
  odataQuery: string | undefined,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  const path = `${aiUserOnlineMeetingBase(userId, onlineMeetingId)}/aiInsights/$count`;
  return graphInvoke(token, { method: 'GET', path: appendCopilotODataQuery(path, odataQuery), beta });
}

export async function copilotMeetingAiInsightPatch(
  token: string,
  userId: string,
  onlineMeetingId: string,
  aiInsightId: string,
  body: Record<string, unknown>,
  beta: boolean,
  extraHeaders?: Record<string, string>
): Promise<GraphResponse<unknown>> {
  const path = `${aiUserOnlineMeetingBase(userId, onlineMeetingId)}/aiInsights/${encodeURIComponent(aiInsightId.trim())}`;
  return graphInvoke(token, { method: 'PATCH', path, body, beta, extraHeaders });
}

export async function copilotMeetingAiInsightDelete(
  token: string,
  userId: string,
  onlineMeetingId: string,
  aiInsightId: string,
  beta: boolean,
  extraHeaders?: Record<string, string>
): Promise<GraphResponse<unknown>> {
  const path = `${aiUserOnlineMeetingBase(userId, onlineMeetingId)}/aiInsights/${encodeURIComponent(aiInsightId.trim())}`;
  return graphInvoke(token, { method: 'DELETE', path, beta, extraHeaders });
}

export async function copilotReportGet(
  token: string,
  fn: CopilotReportFunctionName,
  period: string,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  return graphInvoke(token, { method: 'GET', path: copilotReportPath(fn, period), beta });
}

export async function copilotPackagesList(
  token: string,
  odataQuery: string | undefined
): Promise<GraphResponse<unknown>> {
  let path = '/copilot/admin/catalog/packages';
  if (odataQuery?.trim()) {
    path += odataQuery.trim().startsWith('?') ? odataQuery.trim() : `?${odataQuery.trim()}`;
  }
  return graphInvoke(token, { method: 'GET', path, beta: true });
}

export async function copilotPackagesGet(token: string, packageId: string): Promise<GraphResponse<unknown>> {
  const id = encodeURIComponent(packageId.trim());
  return graphInvoke(token, { method: 'GET', path: `/copilot/admin/catalog/packages/${id}`, beta: true });
}

export async function copilotPackagesUpdate(
  token: string,
  packageId: string,
  body: Record<string, unknown>
): Promise<GraphResponse<unknown>> {
  const id = encodeURIComponent(packageId.trim());
  return graphInvoke(token, { method: 'PATCH', path: `/copilot/admin/catalog/packages/${id}`, body, beta: true });
}

export async function copilotPackagesBlock(token: string, packageId: string): Promise<GraphResponse<unknown>> {
  const id = encodeURIComponent(packageId.trim());
  return graphInvoke(token, { method: 'POST', path: `/copilot/admin/catalog/packages/${id}/block`, beta: true });
}

export async function copilotPackagesUnblock(token: string, packageId: string): Promise<GraphResponse<unknown>> {
  const id = encodeURIComponent(packageId.trim());
  return graphInvoke(token, { method: 'POST', path: `/copilot/admin/catalog/packages/${id}/unblock`, beta: true });
}

export async function copilotPackagesReassign(
  token: string,
  packageId: string,
  newOwnerUserId: string
): Promise<GraphResponse<unknown>> {
  const id = encodeURIComponent(packageId.trim());
  return graphInvoke(token, {
    method: 'POST',
    path: `/copilot/admin/catalog/packages/${id}/reassign`,
    body: { userId: newOwnerUserId.trim() },
    beta: true
  });
}

export async function copilotConversationsList(
  token: string,
  odataQuery: string | undefined,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  return graphInvoke(token, {
    method: 'GET',
    path: appendCopilotODataQuery('/copilot/conversations', odataQuery),
    beta
  });
}

export async function copilotConversationGet(
  token: string,
  conversationId: string,
  odataQuery: string | undefined,
  beta: boolean,
  extraHeaders?: Record<string, string>
): Promise<GraphResponse<unknown>> {
  const id = encodeURIComponent(conversationId.trim());
  return graphInvoke(token, {
    method: 'GET',
    path: appendCopilotODataQuery(`/copilot/conversations/${id}`, odataQuery),
    beta,
    extraHeaders
  });
}

export async function copilotConversationPatch(
  token: string,
  conversationId: string,
  body: Record<string, unknown>,
  beta: boolean,
  extraHeaders?: Record<string, string>
): Promise<GraphResponse<unknown>> {
  const id = encodeURIComponent(conversationId.trim());
  return graphInvoke(token, {
    method: 'PATCH',
    path: `/copilot/conversations/${id}`,
    body,
    beta,
    extraHeaders
  });
}

export async function copilotConversationDelete(
  token: string,
  conversationId: string,
  beta: boolean,
  extraHeaders?: Record<string, string>
): Promise<GraphResponse<unknown>> {
  const id = encodeURIComponent(conversationId.trim());
  return graphInvoke(token, {
    method: 'DELETE',
    path: `/copilot/conversations/${id}`,
    beta,
    extraHeaders
  });
}

export async function copilotConversationDeleteByThreadId(
  token: string,
  body: Record<string, unknown>,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  return graphInvoke(token, {
    method: 'POST',
    path: '/copilot/conversations/microsoft.graph.copilot.deleteByThreadId',
    body,
    beta
  });
}

export async function copilotConversationMessagesList(
  token: string,
  conversationId: string,
  odataQuery: string | undefined,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  const id = encodeURIComponent(conversationId.trim());
  return graphInvoke(token, {
    method: 'GET',
    path: appendCopilotODataQuery(`/copilot/conversations/${id}/messages`, odataQuery),
    beta
  });
}

export async function copilotConversationMessageGet(
  token: string,
  conversationId: string,
  messageId: string,
  odataQuery: string | undefined,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  const c = encodeURIComponent(conversationId.trim());
  const m = encodeURIComponent(messageId.trim());
  return graphInvoke(token, {
    method: 'GET',
    path: appendCopilotODataQuery(`/copilot/conversations/${c}/messages/${m}`, odataQuery),
    beta
  });
}

export async function copilotConversationMessageCreate(
  token: string,
  conversationId: string,
  body: Record<string, unknown>,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  const id = encodeURIComponent(conversationId.trim());
  return graphInvoke(token, {
    method: 'POST',
    path: `/copilot/conversations/${id}/messages`,
    body,
    beta
  });
}

export async function copilotConversationMessagePatch(
  token: string,
  conversationId: string,
  messageId: string,
  body: Record<string, unknown>,
  beta: boolean,
  extraHeaders?: Record<string, string>
): Promise<GraphResponse<unknown>> {
  const c = encodeURIComponent(conversationId.trim());
  const m = encodeURIComponent(messageId.trim());
  return graphInvoke(token, {
    method: 'PATCH',
    path: `/copilot/conversations/${c}/messages/${m}`,
    body,
    beta,
    extraHeaders
  });
}

export async function copilotConversationMessageDelete(
  token: string,
  conversationId: string,
  messageId: string,
  beta: boolean,
  extraHeaders?: Record<string, string>
): Promise<GraphResponse<unknown>> {
  const c = encodeURIComponent(conversationId.trim());
  const m = encodeURIComponent(messageId.trim());
  return graphInvoke(token, {
    method: 'DELETE',
    path: `/copilot/conversations/${c}/messages/${m}`,
    beta,
    extraHeaders
  });
}

export async function copilotAgentsList(
  token: string,
  odataQuery: string | undefined,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  return graphInvoke(token, {
    method: 'GET',
    path: appendCopilotODataQuery('/copilot/agents', odataQuery),
    beta
  });
}

export async function copilotAgentGet(
  token: string,
  agentId: string,
  odataQuery: string | undefined,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  const id = encodeURIComponent(agentId.trim());
  return graphInvoke(token, {
    method: 'GET',
    path: appendCopilotODataQuery(`/copilot/agents/${id}`, odataQuery),
    beta
  });
}

export async function copilotSettingsGet(
  token: string,
  odataQuery: string | undefined,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  return graphInvoke(token, {
    method: 'GET',
    path: appendCopilotODataQuery('/copilot/settings', odataQuery),
    beta
  });
}

export async function copilotSettingsPatch(
  token: string,
  body: Record<string, unknown>,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  return graphInvoke(token, { method: 'PATCH', path: '/copilot/settings', body, beta });
}

export async function copilotSettingsPeopleGet(
  token: string,
  odataQuery: string | undefined,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  return graphInvoke(token, {
    method: 'GET',
    path: appendCopilotODataQuery('/copilot/settings/people', odataQuery),
    beta
  });
}

export async function copilotSettingsPeoplePatch(
  token: string,
  body: Record<string, unknown>,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  return graphInvoke(token, { method: 'PATCH', path: '/copilot/settings/people', body, beta });
}

export async function copilotSettingsEnhancedPersonalizationGet(
  token: string,
  odataQuery: string | undefined,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  return graphInvoke(token, {
    method: 'GET',
    path: appendCopilotODataQuery('/copilot/settings/people/enhancedPersonalization', odataQuery),
    beta
  });
}

export async function copilotSettingsEnhancedPersonalizationPatch(
  token: string,
  body: Record<string, unknown>,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  return graphInvoke(token, {
    method: 'PATCH',
    path: '/copilot/settings/people/enhancedPersonalization',
    body,
    beta
  });
}

export async function copilotSettingsDelete(
  token: string,
  beta: boolean,
  extraHeaders?: Record<string, string>
): Promise<GraphResponse<unknown>> {
  return graphInvoke(token, { method: 'DELETE', path: '/copilot/settings', beta, extraHeaders });
}

export async function copilotSettingsPeopleDelete(
  token: string,
  beta: boolean,
  extraHeaders?: Record<string, string>
): Promise<GraphResponse<unknown>> {
  return graphInvoke(token, { method: 'DELETE', path: '/copilot/settings/people', beta, extraHeaders });
}

export async function copilotSettingsEnhancedPersonalizationDelete(
  token: string,
  beta: boolean,
  extraHeaders?: Record<string, string>
): Promise<GraphResponse<unknown>> {
  return graphInvoke(token, {
    method: 'DELETE',
    path: '/copilot/settings/people/enhancedPersonalization',
    beta,
    extraHeaders
  });
}

export async function copilotReportsNavGet(
  token: string,
  odataQuery: string | undefined,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  return graphInvoke(token, {
    method: 'GET',
    path: appendCopilotODataQuery('/copilot/reports', odataQuery),
    beta
  });
}

export async function copilotReportsNavPatch(
  token: string,
  body: Record<string, unknown>,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  return graphInvoke(token, { method: 'PATCH', path: '/copilot/reports', body, beta });
}

export async function copilotReportsNavDelete(
  token: string,
  beta: boolean,
  extraHeaders?: Record<string, string>
): Promise<GraphResponse<unknown>> {
  return graphInvoke(token, { method: 'DELETE', path: '/copilot/reports', beta, extraHeaders });
}

export async function copilotAdminSettingsGet(
  token: string,
  odataQuery: string | undefined
): Promise<GraphResponse<unknown>> {
  return graphInvoke(token, {
    method: 'GET',
    path: appendCopilotODataQuery('/copilot/admin/settings', odataQuery),
    beta: true
  });
}

export async function copilotAdminSettingsPatch(
  token: string,
  body: Record<string, unknown>
): Promise<GraphResponse<unknown>> {
  return graphInvoke(token, { method: 'PATCH', path: '/copilot/admin/settings', body, beta: true });
}

export async function copilotAdminLimitedModeGet(
  token: string,
  odataQuery: string | undefined
): Promise<GraphResponse<unknown>> {
  return graphInvoke(token, {
    method: 'GET',
    path: appendCopilotODataQuery('/copilot/admin/settings/limitedMode', odataQuery),
    beta: true
  });
}

export async function copilotAdminLimitedModePatch(
  token: string,
  body: Record<string, unknown>
): Promise<GraphResponse<unknown>> {
  return graphInvoke(token, {
    method: 'PATCH',
    path: '/copilot/admin/settings/limitedMode',
    body,
    beta: true
  });
}

export async function copilotPackagesCreate(
  token: string,
  body: Record<string, unknown>
): Promise<GraphResponse<unknown>> {
  return graphInvoke(token, { method: 'POST', path: '/copilot/admin/catalog/packages', body, beta: true });
}

export async function copilotPackagesDelete(
  token: string,
  packageId: string,
  extraHeaders?: Record<string, string>
): Promise<GraphResponse<unknown>> {
  const id = encodeURIComponent(packageId.trim());
  return graphInvoke(token, {
    method: 'DELETE',
    path: `/copilot/admin/catalog/packages/${id}`,
    beta: true,
    extraHeaders
  });
}

/** GET package zip bytes (beta). */
export async function copilotPackageZipDownload(token: string, packageId: string): Promise<GraphResponse<Uint8Array>> {
  const id = encodeURIComponent(packageId.trim());
  const path = `/copilot/admin/catalog/packages/${id}/zipFile`;
  const url = `${GRAPH_BETA_URL.replace(/\/$/, '')}${path}`;
  try {
    // codeql[js/file-access-to-http]: Bearer token from caller; URL is fixed Graph beta package zip path.
    const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
    if (!res.ok) {
      let message = `HTTP ${res.status}`;
      try {
        const json = (await res.json()) as { error?: { message?: string } };
        message = json.error?.message || message;
      } catch {
        // ignore
      }
      return graphError(message, undefined, res.status);
    }
    return graphResult(new Uint8Array(await res.arrayBuffer()));
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Download failed');
  }
}

/** PUT package zip (beta). */
export async function copilotPackageZipUpload(
  token: string,
  packageId: string,
  bytes: Uint8Array,
  contentType: string
): Promise<GraphResponse<unknown>> {
  const id = encodeURIComponent(packageId.trim());
  const path = `/copilot/admin/catalog/packages/${id}/zipFile`;
  const url = `${GRAPH_BETA_URL.replace(/\/$/, '')}${path}`;
  try {
    // codeql[js/file-access-to-http]: Bearer token from caller; URL is fixed Graph beta package zip path.
    const res = await fetch(url, {
      method: 'PUT',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': contentType || 'application/octet-stream'
      },
      body: new Blob([Uint8Array.from(bytes)])
    });
    if (!res.ok) {
      let message = `HTTP ${res.status}`;
      try {
        const json = (await res.json()) as { error?: { message?: string } };
        message = json.error?.message || message;
      } catch {
        // ignore
      }
      return graphError(message, undefined, res.status);
    }
    if (res.status === 204) return graphResult(undefined);
    const ct = res.headers.get('content-type') || '';
    if (ct.includes('application/json')) {
      return graphResult((await res.json()) as unknown);
    }
    return graphResult(undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Upload failed');
  }
}

export async function copilotRealtimeActivityFeedGet(
  token: string,
  odataQuery: string | undefined,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  return graphInvoke(token, {
    method: 'GET',
    path: appendCopilotODataQuery(REALTIME_ACTIVITY, odataQuery),
    beta
  });
}

export async function copilotRealtimeMeetingsList(
  token: string,
  odataQuery: string | undefined,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  return graphInvoke(token, {
    method: 'GET',
    path: appendCopilotODataQuery(`${REALTIME_ACTIVITY}/meetings`, odataQuery),
    beta
  });
}

export async function copilotRealtimeMeetingCreate(
  token: string,
  body: Record<string, unknown>,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  return graphInvoke(token, {
    method: 'POST',
    path: `${REALTIME_ACTIVITY}/meetings`,
    body,
    beta
  });
}

export async function copilotRealtimeMeetingGet(
  token: string,
  meetingId: string,
  odataQuery: string | undefined,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  const id = encodeURIComponent(meetingId.trim());
  return graphInvoke(token, {
    method: 'GET',
    path: appendCopilotODataQuery(`${REALTIME_ACTIVITY}/meetings/${id}`, odataQuery),
    beta
  });
}

export async function copilotRealtimeMeetingPatch(
  token: string,
  meetingId: string,
  body: Record<string, unknown>,
  beta: boolean,
  extraHeaders?: Record<string, string>
): Promise<GraphResponse<unknown>> {
  const id = encodeURIComponent(meetingId.trim());
  return graphInvoke(token, {
    method: 'PATCH',
    path: `${REALTIME_ACTIVITY}/meetings/${id}`,
    body,
    beta,
    extraHeaders
  });
}

export async function copilotRealtimeMeetingDelete(
  token: string,
  meetingId: string,
  beta: boolean,
  extraHeaders?: Record<string, string>
): Promise<GraphResponse<unknown>> {
  const id = encodeURIComponent(meetingId.trim());
  return graphInvoke(token, {
    method: 'DELETE',
    path: `${REALTIME_ACTIVITY}/meetings/${id}`,
    beta,
    extraHeaders
  });
}

export async function copilotRealtimeSubscriptionsList(
  token: string,
  odataQuery: string | undefined,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  return graphInvoke(token, {
    method: 'GET',
    path: appendCopilotODataQuery(`${REALTIME_ACTIVITY}/multiActivitySubscriptions`, odataQuery),
    beta
  });
}

export async function copilotRealtimeSubscriptionCreate(
  token: string,
  body: Record<string, unknown>,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  return graphInvoke(token, {
    method: 'POST',
    path: `${REALTIME_ACTIVITY}/multiActivitySubscriptions`,
    body,
    beta
  });
}

export async function copilotRealtimeSubscriptionGet(
  token: string,
  subscriptionId: string,
  odataQuery: string | undefined,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  const id = encodeURIComponent(subscriptionId.trim());
  return graphInvoke(token, {
    method: 'GET',
    path: appendCopilotODataQuery(`${REALTIME_ACTIVITY}/multiActivitySubscriptions/${id}`, odataQuery),
    beta
  });
}

export async function copilotRealtimeSubscriptionPatch(
  token: string,
  subscriptionId: string,
  body: Record<string, unknown>,
  beta: boolean,
  extraHeaders?: Record<string, string>
): Promise<GraphResponse<unknown>> {
  const id = encodeURIComponent(subscriptionId.trim());
  return graphInvoke(token, {
    method: 'PATCH',
    path: `${REALTIME_ACTIVITY}/multiActivitySubscriptions/${id}`,
    body,
    beta,
    extraHeaders
  });
}

export async function copilotRealtimeSubscriptionDelete(
  token: string,
  subscriptionId: string,
  beta: boolean,
  extraHeaders?: Record<string, string>
): Promise<GraphResponse<unknown>> {
  const id = encodeURIComponent(subscriptionId.trim());
  return graphInvoke(token, {
    method: 'DELETE',
    path: `${REALTIME_ACTIVITY}/multiActivitySubscriptions/${id}`,
    beta,
    extraHeaders
  });
}

export async function copilotRealtimeSubscriptionGetArtifacts(
  token: string,
  subscriptionId: string,
  body: Record<string, unknown> | undefined,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  const id = encodeURIComponent(subscriptionId.trim());
  return graphInvoke(token, {
    method: 'POST',
    path: `${REALTIME_ACTIVITY}/multiActivitySubscriptions/${id}/getArtifacts`,
    body: body ?? {},
    beta
  });
}

export async function copilotRealtimeTranscriptsList(
  token: string,
  meetingId: string,
  odataQuery: string | undefined,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  const m = encodeURIComponent(meetingId.trim());
  return graphInvoke(token, {
    method: 'GET',
    path: appendCopilotODataQuery(`${REALTIME_ACTIVITY}/meetings/${m}/transcripts`, odataQuery),
    beta
  });
}

export async function copilotRealtimeTranscriptCreate(
  token: string,
  meetingId: string,
  body: Record<string, unknown>,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  const m = encodeURIComponent(meetingId.trim());
  return graphInvoke(token, {
    method: 'POST',
    path: `${REALTIME_ACTIVITY}/meetings/${m}/transcripts`,
    body,
    beta
  });
}

export async function copilotRealtimeTranscriptGet(
  token: string,
  meetingId: string,
  transcriptId: string,
  odataQuery: string | undefined,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  const m = encodeURIComponent(meetingId.trim());
  const t = encodeURIComponent(transcriptId.trim());
  return graphInvoke(token, {
    method: 'GET',
    path: appendCopilotODataQuery(`${REALTIME_ACTIVITY}/meetings/${m}/transcripts/${t}`, odataQuery),
    beta
  });
}

export async function copilotRealtimeTranscriptPatch(
  token: string,
  meetingId: string,
  transcriptId: string,
  body: Record<string, unknown>,
  beta: boolean,
  extraHeaders?: Record<string, string>
): Promise<GraphResponse<unknown>> {
  const m = encodeURIComponent(meetingId.trim());
  const t = encodeURIComponent(transcriptId.trim());
  return graphInvoke(token, {
    method: 'PATCH',
    path: `${REALTIME_ACTIVITY}/meetings/${m}/transcripts/${t}`,
    body,
    beta,
    extraHeaders
  });
}

export async function copilotRealtimeTranscriptDelete(
  token: string,
  meetingId: string,
  transcriptId: string,
  beta: boolean,
  extraHeaders?: Record<string, string>
): Promise<GraphResponse<unknown>> {
  const m = encodeURIComponent(meetingId.trim());
  const t = encodeURIComponent(transcriptId.trim());
  return graphInvoke(token, {
    method: 'DELETE',
    path: `${REALTIME_ACTIVITY}/meetings/${m}/transcripts/${t}`,
    beta,
    extraHeaders
  });
}

/** GET /copilot — Copilot root resource. */
export async function copilotRootGet(
  token: string,
  odataQuery: string | undefined,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  return graphInvoke(token, { method: 'GET', path: appendCopilotODataQuery('/copilot', odataQuery), beta });
}

export async function copilotRootPatch(
  token: string,
  body: Record<string, unknown>,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  return graphInvoke(token, { method: 'PATCH', path: '/copilot', body, beta });
}

export async function copilotAdminNavGet(
  token: string,
  odataQuery: string | undefined
): Promise<GraphResponse<unknown>> {
  return graphInvoke(token, {
    method: 'GET',
    path: appendCopilotODataQuery('/copilot/admin', odataQuery),
    beta: true
  });
}

export async function copilotAdminNavPatch(
  token: string,
  body: Record<string, unknown>
): Promise<GraphResponse<unknown>> {
  return graphInvoke(token, { method: 'PATCH', path: '/copilot/admin', body, beta: true });
}

export async function copilotAdminNavDelete(
  token: string,
  extraHeaders?: Record<string, string>
): Promise<GraphResponse<unknown>> {
  return graphInvoke(token, { method: 'DELETE', path: '/copilot/admin', beta: true, extraHeaders });
}

export async function copilotAdminCatalogGet(
  token: string,
  odataQuery: string | undefined
): Promise<GraphResponse<unknown>> {
  return graphInvoke(token, {
    method: 'GET',
    path: appendCopilotODataQuery('/copilot/admin/catalog', odataQuery),
    beta: true
  });
}

export async function copilotAdminCatalogPatch(
  token: string,
  body: Record<string, unknown>
): Promise<GraphResponse<unknown>> {
  return graphInvoke(token, { method: 'PATCH', path: '/copilot/admin/catalog', body, beta: true });
}

export async function copilotAdminCatalogDelete(
  token: string,
  extraHeaders?: Record<string, string>
): Promise<GraphResponse<unknown>> {
  return graphInvoke(token, { method: 'DELETE', path: '/copilot/admin/catalog', beta: true, extraHeaders });
}

export async function copilotPackagesCount(
  token: string,
  odataQuery: string | undefined
): Promise<GraphResponse<unknown>> {
  return graphInvoke(token, {
    method: 'GET',
    path: appendCopilotODataQuery('/copilot/admin/catalog/packages/$count', odataQuery),
    beta: true
  });
}

export async function copilotPackageZipDelete(
  token: string,
  packageId: string,
  extraHeaders?: Record<string, string>
): Promise<GraphResponse<unknown>> {
  const id = encodeURIComponent(packageId.trim());
  return graphInvoke(token, {
    method: 'DELETE',
    path: `/copilot/admin/catalog/packages/${id}/zipFile`,
    beta: true,
    extraHeaders
  });
}

export async function copilotAdminSettingsDelete(
  token: string,
  extraHeaders?: Record<string, string>
): Promise<GraphResponse<unknown>> {
  return graphInvoke(token, { method: 'DELETE', path: '/copilot/admin/settings', beta: true, extraHeaders });
}

export async function copilotAdminLimitedModeDelete(
  token: string,
  extraHeaders?: Record<string, string>
): Promise<GraphResponse<unknown>> {
  return graphInvoke(token, {
    method: 'DELETE',
    path: '/copilot/admin/settings/limitedMode',
    beta: true,
    extraHeaders
  });
}

export async function copilotCommunicationsGet(
  token: string,
  odataQuery: string | undefined,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  return graphInvoke(token, {
    method: 'GET',
    path: appendCopilotODataQuery(COPILOT_COMMUNICATIONS, odataQuery),
    beta
  });
}

export async function copilotCommunicationsPatch(
  token: string,
  body: Record<string, unknown>,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  return graphInvoke(token, { method: 'PATCH', path: COPILOT_COMMUNICATIONS, body, beta });
}

export async function copilotCommunicationsDelete(
  token: string,
  beta: boolean,
  extraHeaders?: Record<string, string>
): Promise<GraphResponse<unknown>> {
  return graphInvoke(token, { method: 'DELETE', path: COPILOT_COMMUNICATIONS, beta, extraHeaders });
}

export async function copilotRealtimeActivityFeedPatch(
  token: string,
  body: Record<string, unknown>,
  beta: boolean,
  extraHeaders?: Record<string, string>
): Promise<GraphResponse<unknown>> {
  return graphInvoke(token, {
    method: 'PATCH',
    path: REALTIME_ACTIVITY,
    body,
    beta,
    extraHeaders
  });
}

export async function copilotRealtimeActivityFeedDelete(
  token: string,
  beta: boolean,
  extraHeaders?: Record<string, string>
): Promise<GraphResponse<unknown>> {
  return graphInvoke(token, { method: 'DELETE', path: REALTIME_ACTIVITY, beta, extraHeaders });
}

export async function copilotConversationsCount(
  token: string,
  odataQuery: string | undefined,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  return graphInvoke(token, {
    method: 'GET',
    path: appendCopilotODataQuery('/copilot/conversations/$count', odataQuery),
    beta
  });
}

export async function copilotConversationMessagesCount(
  token: string,
  conversationId: string,
  odataQuery: string | undefined,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  const c = encodeURIComponent(conversationId.trim());
  return graphInvoke(token, {
    method: 'GET',
    path: appendCopilotODataQuery(`/copilot/conversations/${c}/messages/$count`, odataQuery),
    beta
  });
}

export async function copilotAgentsCount(
  token: string,
  odataQuery: string | undefined,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  return graphInvoke(token, {
    method: 'GET',
    path: appendCopilotODataQuery('/copilot/agents/$count', odataQuery),
    beta
  });
}

export async function copilotRealtimeMeetingsCount(
  token: string,
  odataQuery: string | undefined,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  return graphInvoke(token, {
    method: 'GET',
    path: appendCopilotODataQuery(`${REALTIME_ACTIVITY}/meetings/$count`, odataQuery),
    beta
  });
}

export async function copilotRealtimeTranscriptsCount(
  token: string,
  meetingId: string,
  odataQuery: string | undefined,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  const m = encodeURIComponent(meetingId.trim());
  return graphInvoke(token, {
    method: 'GET',
    path: appendCopilotODataQuery(`${REALTIME_ACTIVITY}/meetings/${m}/transcripts/$count`, odataQuery),
    beta
  });
}

export async function copilotRealtimeSubscriptionsCount(
  token: string,
  odataQuery: string | undefined,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  return graphInvoke(token, {
    method: 'GET',
    path: appendCopilotODataQuery(`${REALTIME_ACTIVITY}/multiActivitySubscriptions/$count`, odataQuery),
    beta
  });
}

export async function copilotInteractionHistoryNavGet(
  token: string,
  odataQuery: string | undefined,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  return graphInvoke(token, {
    method: 'GET',
    path: appendCopilotODataQuery('/copilot/interactionHistory', odataQuery),
    beta
  });
}

export async function copilotInteractionHistoryNavPatch(
  token: string,
  body: Record<string, unknown>,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  return graphInvoke(token, { method: 'PATCH', path: '/copilot/interactionHistory', body, beta });
}

export async function copilotInteractionHistoryNavDelete(
  token: string,
  beta: boolean,
  extraHeaders?: Record<string, string>
): Promise<GraphResponse<unknown>> {
  return graphInvoke(token, { method: 'DELETE', path: '/copilot/interactionHistory', beta, extraHeaders });
}

export async function copilotAiUsersList(
  token: string,
  odataQuery: string | undefined,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  return graphInvoke(token, { method: 'GET', path: appendCopilotODataQuery('/copilot/users', odataQuery), beta });
}

export async function copilotAiUsersCount(
  token: string,
  odataQuery: string | undefined,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  return graphInvoke(token, {
    method: 'GET',
    path: appendCopilotODataQuery('/copilot/users/$count', odataQuery),
    beta
  });
}

export async function copilotAiUserCreate(
  token: string,
  body: Record<string, unknown>,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  return graphInvoke(token, { method: 'POST', path: '/copilot/users', body, beta });
}

export async function copilotAiUserGet(
  token: string,
  aiUserId: string,
  odataQuery: string | undefined,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  const id = encodeURIComponent(aiUserId.trim());
  return graphInvoke(token, {
    method: 'GET',
    path: appendCopilotODataQuery(`/copilot/users/${id}`, odataQuery),
    beta
  });
}

export async function copilotAiUserPatch(
  token: string,
  aiUserId: string,
  body: Record<string, unknown>,
  beta: boolean,
  extraHeaders?: Record<string, string>
): Promise<GraphResponse<unknown>> {
  const id = encodeURIComponent(aiUserId.trim());
  return graphInvoke(token, { method: 'PATCH', path: `/copilot/users/${id}`, body, beta, extraHeaders });
}

export async function copilotAiUserDelete(
  token: string,
  aiUserId: string,
  beta: boolean,
  extraHeaders?: Record<string, string>
): Promise<GraphResponse<unknown>> {
  const id = encodeURIComponent(aiUserId.trim());
  return graphInvoke(token, { method: 'DELETE', path: `/copilot/users/${id}`, beta, extraHeaders });
}

export async function copilotAiUserInteractionHistoryGet(
  token: string,
  aiUserId: string,
  odataQuery: string | undefined,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  const u = copilotUserSegment(aiUserId);
  return graphInvoke(token, {
    method: 'GET',
    path: appendCopilotODataQuery(`/copilot/users/${u}/interactionHistory`, odataQuery),
    beta
  });
}

export async function copilotAiUserInteractionHistoryPatch(
  token: string,
  aiUserId: string,
  body: Record<string, unknown>,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  const u = copilotUserSegment(aiUserId);
  return graphInvoke(token, {
    method: 'PATCH',
    path: `/copilot/users/${u}/interactionHistory`,
    body,
    beta
  });
}

export async function copilotAiUserInteractionHistoryDelete(
  token: string,
  aiUserId: string,
  beta: boolean,
  extraHeaders?: Record<string, string>
): Promise<GraphResponse<unknown>> {
  const u = copilotUserSegment(aiUserId);
  return graphInvoke(token, {
    method: 'DELETE',
    path: `/copilot/users/${u}/interactionHistory`,
    beta,
    extraHeaders
  });
}

export async function copilotAiUserOnlineMeetingsList(
  token: string,
  aiUserId: string,
  odataQuery: string | undefined,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  const u = copilotUserSegment(aiUserId);
  return graphInvoke(token, {
    method: 'GET',
    path: appendCopilotODataQuery(`/copilot/users/${u}/onlineMeetings`, odataQuery),
    beta
  });
}

export async function copilotAiUserOnlineMeetingsCount(
  token: string,
  aiUserId: string,
  odataQuery: string | undefined,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  const u = copilotUserSegment(aiUserId);
  return graphInvoke(token, {
    method: 'GET',
    path: appendCopilotODataQuery(`/copilot/users/${u}/onlineMeetings/$count`, odataQuery),
    beta
  });
}

export async function copilotAiUserOnlineMeetingCreate(
  token: string,
  aiUserId: string,
  body: Record<string, unknown>,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  const u = copilotUserSegment(aiUserId);
  return graphInvoke(token, { method: 'POST', path: `/copilot/users/${u}/onlineMeetings`, body, beta });
}

export async function copilotAiUserOnlineMeetingGet(
  token: string,
  aiUserId: string,
  onlineMeetingId: string,
  odataQuery: string | undefined,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  const base = aiUserOnlineMeetingBase(aiUserId, onlineMeetingId);
  return graphInvoke(token, { method: 'GET', path: appendCopilotODataQuery(base, odataQuery), beta });
}

export async function copilotAiUserOnlineMeetingPatch(
  token: string,
  aiUserId: string,
  onlineMeetingId: string,
  body: Record<string, unknown>,
  beta: boolean,
  extraHeaders?: Record<string, string>
): Promise<GraphResponse<unknown>> {
  const path = aiUserOnlineMeetingBase(aiUserId, onlineMeetingId);
  return graphInvoke(token, { method: 'PATCH', path, body, beta, extraHeaders });
}

export async function copilotAiUserOnlineMeetingDelete(
  token: string,
  aiUserId: string,
  onlineMeetingId: string,
  beta: boolean,
  extraHeaders?: Record<string, string>
): Promise<GraphResponse<unknown>> {
  const path = aiUserOnlineMeetingBase(aiUserId, onlineMeetingId);
  return graphInvoke(token, { method: 'DELETE', path, beta, extraHeaders });
}
