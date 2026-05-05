/**
 * Microsoft 365 Copilot APIs on Microsoft Graph (`/copilot/...`).
 * @see https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/copilot-apis-overview
 */

import { callGraphAbsolute, GraphApiError, type GraphResponse } from './graph-client.js';
import { graphInvoke, graphInvokeText } from './graph-advanced-client.js';

export const COPILOT_REPORT_PERIODS = ['D7', 'D30', 'D90', 'D180', 'ALL'] as const;
export type CopilotReportPeriod = (typeof COPILOT_REPORT_PERIODS)[number];

const MAX_SEARCH_QUERY = 1500;
const MAX_RETRIEVAL_QUERY = 1500;

export const COPILOT_RETRIEVAL_DATA_SOURCES = ['sharePoint', 'oneDriveBusiness', 'externalItem'] as const;
export type CopilotRetrievalDataSource = (typeof COPILOT_RETRIEVAL_DATA_SOURCES)[number];

export function assertCopilotRetrievalDataSource(s: string): CopilotRetrievalDataSource {
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
    throw new GraphApiError(
      `period must be one of: ${COPILOT_REPORT_PERIODS.join(', ')}`,
      'InvalidPeriod',
      400
    );
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

export async function copilotConversationCreate(
  token: string,
  beta: boolean = true
): Promise<GraphResponse<unknown>> {
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
    path: `/copilot/conversations/${id}/chat`,
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
    path: `/copilot/conversations/${id}/chatOverStream`,
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
  const base = `/copilot/users/${copilotUserSegment(userId)}/interactionHistory/getAllEnterpriseInteractions`;
  const q = rawODataQuery?.trim();
  const path = q ? `${base}?${q.replace(/^\?/, '')}` : base;
  return graphInvoke(token, { method: 'GET', path, beta });
}

export async function copilotMeetingInsightsList(
  token: string,
  userId: string,
  onlineMeetingId: string,
  odataQuery: string | undefined,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  const u = copilotUserSegment(userId);
  const m = encodeURIComponent(onlineMeetingId.trim());
  let path = `/copilot/users/${u}/onlineMeetings/${m}/aiInsights`;
  if (odataQuery?.trim()) {
    path += odataQuery.trim().startsWith('?') ? odataQuery.trim() : `?${odataQuery.trim()}`;
  }
  return graphInvoke(token, { method: 'GET', path, beta });
}

export async function copilotMeetingInsightGet(
  token: string,
  userId: string,
  onlineMeetingId: string,
  aiInsightId: string,
  odataQuery: string | undefined,
  beta: boolean
): Promise<GraphResponse<unknown>> {
  const u = copilotUserSegment(userId);
  const m = encodeURIComponent(onlineMeetingId.trim());
  const i = encodeURIComponent(aiInsightId.trim());
  let path = `/copilot/users/${u}/onlineMeetings/${m}/aiInsights/${i}`;
  if (odataQuery?.trim()) {
    path += odataQuery.trim().startsWith('?') ? odataQuery.trim() : `?${odataQuery.trim()}`;
  }
  return graphInvoke(token, { method: 'GET', path, beta });
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
