import {
  callGraph,
  fetchAllPages,
  GraphApiError,
  type GraphResponse,
  graphError,
  graphResult
} from './graph-client.js';

export interface PlannerPlan {
  id: string;
  title: string;
  owner?: string;
}

export interface PlannerBucket {
  id: string;
  name: string;
  planId: string;
  orderHint?: string;
}

/** Planner label slots (plan defines display names in plan details). */
export type PlannerCategorySlot = 'category1' | 'category2' | 'category3' | 'category4' | 'category5' | 'category6';

export type PlannerAppliedCategories = Partial<Record<PlannerCategorySlot, boolean>>;

export interface PlannerTask {
  id: string;
  planId: string;
  bucketId: string;
  title: string;
  orderHint: string;
  assigneePriority: string;
  percentComplete: number;
  hasDescription: boolean;
  createdDateTime: string;
  dueDateTime?: string;
  assignments?: Record<string, any>;
  /** Label slots (boolean per category1..category6); names come from plan details. */
  appliedCategories?: PlannerAppliedCategories;
  '@odata.etag'?: string;
}

export interface PlannerPlanDetails {
  id: string;
  categoryDescriptions?: Partial<Record<PlannerCategorySlot, string>>;
}

const PLANNER_SLOTS: PlannerCategorySlot[] = [
  'category1',
  'category2',
  'category3',
  'category4',
  'category5',
  'category6'
];

/** Accept `1`..`6` or `category1`..`category6` (case-insensitive). */
export function parsePlannerLabelKey(input: string): PlannerCategorySlot | null {
  const t = input.trim().toLowerCase();
  const m = t.match(/^category([1-6])$/);
  if (m) return `category${m[1]}` as PlannerCategorySlot;
  if (/^[1-6]$/.test(t)) return `category${t}` as PlannerCategorySlot;
  return null;
}

/** Build a full slot map for PATCH (Planner expects explicit booleans per slot). */
export function normalizeAppliedCategories(
  current: PlannerAppliedCategories | undefined,
  patch: { clearAll?: boolean; setTrue?: PlannerCategorySlot[]; setFalse?: PlannerCategorySlot[] }
): PlannerAppliedCategories {
  const out: PlannerAppliedCategories = {};
  for (const s of PLANNER_SLOTS) {
    if (patch.clearAll) {
      out[s] = false;
      continue;
    }
    let v = current?.[s] === true;
    for (const u of patch.setTrue ?? []) if (u === s) v = true;
    for (const u of patch.setFalse ?? []) if (u === s) v = false;
    out[s] = v;
  }
  return out;
}

export async function listUserTasks(token: string): Promise<GraphResponse<PlannerTask[]>> {
  return fetchAllPages<PlannerTask>(token, '/me/planner/tasks', 'Failed to list tasks');
}

export async function listUserPlans(token: string): Promise<GraphResponse<PlannerPlan[]>> {
  return fetchAllPages<PlannerPlan>(token, '/me/planner/plans', 'Failed to list plans');
}

export async function listGroupPlans(token: string, groupId: string): Promise<GraphResponse<PlannerPlan[]>> {
  return fetchAllPages<PlannerPlan>(
    token,
    `/groups/${encodeURIComponent(groupId)}/planner/plans`,
    'Failed to list group plans'
  );
}

export async function listPlanBuckets(token: string, planId: string): Promise<GraphResponse<PlannerBucket[]>> {
  return fetchAllPages<PlannerBucket>(
    token,
    `/planner/plans/${encodeURIComponent(planId)}/buckets`,
    'Failed to list buckets'
  );
}

export async function listPlanTasks(token: string, planId: string): Promise<GraphResponse<PlannerTask[]>> {
  return fetchAllPages<PlannerTask>(
    token,
    `/planner/plans/${encodeURIComponent(planId)}/tasks`,
    'Failed to list plan tasks'
  );
}

export async function getPlanDetails(token: string, planId: string): Promise<GraphResponse<PlannerPlanDetails>> {
  try {
    const result = await callGraph<PlannerPlanDetails>(token, `/planner/plans/${encodeURIComponent(planId)}/details`);
    if (!result.ok || !result.data) {
      return graphError(
        result.error?.message || 'Failed to get plan details',
        result.error?.code,
        result.error?.status
      );
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get plan details');
  }
}

export async function getTask(token: string, taskId: string): Promise<GraphResponse<PlannerTask>> {
  try {
    const result = await callGraph<PlannerTask>(token, `/planner/tasks/${encodeURIComponent(taskId)}`);
    if (!result.ok || !result.data) {
      return graphError(result.error?.message || 'Failed to get task', result.error?.code, result.error?.status);
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to get task');
  }
}

export async function createTask(
  token: string,
  planId: string,
  title: string,
  bucketId?: string,
  assignments?: Record<string, any>,
  appliedCategories?: PlannerAppliedCategories
): Promise<GraphResponse<PlannerTask>> {
  try {
    const body: Record<string, unknown> = { planId, title };
    if (bucketId) body.bucketId = bucketId;
    if (assignments) body.assignments = assignments;

    const result = await callGraph<PlannerTask>(token, '/planner/tasks', {
      method: 'POST',
      body: JSON.stringify(body)
    });
    if (!result.ok || !result.data) {
      return graphError(result.error?.message || 'Failed to create task', result.error?.code, result.error?.status);
    }
    if (appliedCategories && Object.keys(appliedCategories).length > 0) {
      const etag = result.data['@odata.etag'];
      if (!etag) {
        return graphError('Created task missing ETag; cannot set labels', 'MISSING_ETAG', 500);
      }
      const patch = await callGraph<void>(token, `/planner/tasks/${encodeURIComponent(result.data.id)}`, {
        method: 'PATCH',
        headers: { 'If-Match': etag },
        body: JSON.stringify({ appliedCategories })
      });
      if (!patch.ok) {
        return graphError(patch.error?.message || 'Failed to set task labels', patch.error?.code, patch.error?.status);
      }
      const again = await getTask(token, result.data.id);
      if (again.ok && again.data) return graphResult(again.data);
    }
    return graphResult(result.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to create task');
  }
}

export async function updateTask(
  token: string,
  taskId: string,
  etag: string,
  updates: {
    title?: string;
    bucketId?: string;
    assignments?: Record<string, any>;
    percentComplete?: number;
    appliedCategories?: PlannerAppliedCategories;
  }
): Promise<GraphResponse<void>> {
  try {
    const result = await callGraph<void>(token, `/planner/tasks/${encodeURIComponent(taskId)}`, {
      method: 'PATCH',
      headers: {
        'If-Match': etag
      },
      body: JSON.stringify(updates)
    });
    if (!result.ok) {
      return graphError(result.error?.message || 'Failed to update task', result.error?.code, result.error?.status);
    }
    return graphResult(undefined as undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Failed to update task');
  }
}
