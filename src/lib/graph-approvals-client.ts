import {
  callGraphAt,
  GraphApiError,
  type GraphResponse,
  graphError,
  graphErrorFromApiError,
  graphResult
} from './graph-client.js';
import { GRAPH_BETA_URL } from './graph-constants.js';

/** `microsoft.graph.approval` — Teams Approvals + Power Automate approvals (beta). */
export interface Approval {
  id: string;
  /** Steps live as a navigation property; included via `$expand=steps` if requested. */
  steps?: ApprovalStep[];
}

/** `microsoft.graph.approvalStep`. Writable: `reviewResult`, `justification`. */
export interface ApprovalStep {
  id: string;
  displayName?: string;
  assignedToMe?: boolean;
  status?: string;
  reviewResult?: string;
  justification?: string;
  reviewedDateTime?: string;
  reviewedBy?: Array<{ user?: { id?: string; displayName?: string } }>;
}

export interface ApprovalListResponse {
  value?: Approval[];
  '@odata.nextLink'?: string;
}

export interface ApprovalStepListResponse {
  value?: ApprovalStep[];
  '@odata.nextLink'?: string;
}

/**
 * `GET /me/approvals` (beta). Defaults to `$expand=steps` so the caller can render
 * approve/deny actionable items without a second round-trip.
 */
export async function listMyApprovals(
  token: string,
  options: { top?: number; expandSteps?: boolean } = {}
): Promise<GraphResponse<ApprovalListResponse>> {
  const expand = options.expandSteps !== false ? '$expand=steps' : '';
  const top = options.top && options.top > 0 ? `$top=${Math.min(Math.max(1, options.top), 200)}` : '';
  const qs = [expand, top].filter(Boolean).join('&');
  const path = `/me/approvals${qs ? `?${qs}` : ''}`;
  try {
    return await callGraphAt<ApprovalListResponse>(GRAPH_BETA_URL, token, path);
  } catch (err) {
    if (err instanceof GraphApiError) return graphErrorFromApiError(err);
    return graphError(err instanceof Error ? err.message : 'Failed to list /me/approvals');
  }
}

/** `GET /me/approvals/{id}` (beta). */
export async function getApproval(
  token: string,
  approvalId: string,
  options: { expandSteps?: boolean } = {}
): Promise<GraphResponse<Approval>> {
  const expand = options.expandSteps !== false ? '?$expand=steps' : '';
  const path = `/me/approvals/${encodeURIComponent(approvalId)}${expand}`;
  try {
    return await callGraphAt<Approval>(GRAPH_BETA_URL, token, path);
  } catch (err) {
    if (err instanceof GraphApiError) return graphErrorFromApiError(err);
    return graphError(err instanceof Error ? err.message : 'Failed to get approval');
  }
}

/** `GET /me/approvals/{id}/steps` (beta). */
export async function listApprovalSteps(
  token: string,
  approvalId: string
): Promise<GraphResponse<ApprovalStepListResponse>> {
  const path = `/me/approvals/${encodeURIComponent(approvalId)}/steps`;
  try {
    return await callGraphAt<ApprovalStepListResponse>(GRAPH_BETA_URL, token, path);
  } catch (err) {
    if (err instanceof GraphApiError) return graphErrorFromApiError(err);
    return graphError(err instanceof Error ? err.message : 'Failed to list approval steps');
  }
}

/**
 * `PATCH /me/approvals/{id}/steps/{stepId}` (beta) — apply approve or deny decision.
 * `reviewResult`: `"Approve"` | `"Deny"` (Microsoft Graph is case-insensitive but uses TitleCase in docs).
 */
export async function patchApprovalStep(
  token: string,
  approvalId: string,
  stepId: string,
  body: { reviewResult: 'Approve' | 'Deny'; justification?: string }
): Promise<GraphResponse<ApprovalStep>> {
  const path = `/me/approvals/${encodeURIComponent(approvalId)}/steps/${encodeURIComponent(stepId)}`;
  try {
    const r = await callGraphAt<ApprovalStep>(GRAPH_BETA_URL, token, path, {
      method: 'PATCH',
      body: JSON.stringify({
        reviewResult: body.reviewResult,
        ...(body.justification ? { justification: body.justification } : {})
      })
    });
    if (!r.ok) {
      return graphError(r.error?.message || 'Approval step PATCH failed', r.error?.code, r.error?.status);
    }
    // Graph returns 204 No Content on success — `callGraphAt` yields ok with no body.
    if (r.data) {
      return graphResult(r.data);
    }
    return graphResult({
      id: stepId,
      reviewResult: body.reviewResult,
      ...(body.justification ? { justification: body.justification } : {})
    });
  } catch (err) {
    if (err instanceof GraphApiError) return graphErrorFromApiError(err);
    return graphError(err instanceof Error ? err.message : 'Failed to PATCH approval step');
  }
}
