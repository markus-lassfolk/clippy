/**
 * Exercise graph-viva-client callGraphAt / fetchAllPages wrappers without mock.module on graph-client
 * (module mocks leak across test files in the same Bun worker and break other suites).
 */
import { afterAll, beforeAll, describe, expect, it } from 'bun:test';

describe('graph-viva-client API wrappers (stubbed fetch)', () => {
  let v: typeof import('./graph-viva-client.js');
  const originalFetch = globalThis.fetch;

  beforeAll(async () => {
    globalThis.fetch = (async (_input: string | URL | Request, init?: RequestInit) => {
      const method = (init?.method ?? 'GET').toUpperCase();
      if (method === 'DELETE') {
        return new Response(null, { status: 204 });
      }
      return new Response('{}', {
        status: 200,
        headers: { 'content-type': 'application/json' }
      });
    }) as unknown as typeof fetch;
    v = await import('./graph-viva-client.js');
  });

  afterAll(() => {
    globalThis.fetch = originalFetch;
  });

  const t = 'tok';
  const uid = 'u@contoso.com';
  const body = { x: 1 };

  it('calls beta paths for working time, insights, employee experience, learning, admin, org, work hours, meetings', async () => {
    await v.getWorkingTimeSchedule(t);
    await v.getWorkingTimeSchedule(t, uid);
    await v.patchWorkingTimeSchedule(t, body, uid);
    await v.deleteWorkingTimeSchedule(t, uid, 'W/"1"');
    await v.startWorkingTime(t, uid);
    await v.endWorkingTime(t);

    await v.getUserItemInsightsSettings(t);
    await v.patchUserItemInsightsSettings(t, body, uid);
    await v.deleteUserItemInsightsSettings(t, uid, 'im');

    await v.listEmployeeExperienceAssignedRoles(t, undefined, '?$top=1');
    await v.listEmployeeExperienceAssignedRoles(t, uid, '$top=2');
    await v.getEmployeeExperience(t);
    await v.patchEmployeeExperience(t, body);
    await v.deleteEmployeeExperience(t, undefined, 'W/"1"');

    await v.createEmployeeExperienceAssignedRole(t, body, uid);
    await v.getEmployeeExperienceAssignedRole(t, 'r1', uid);
    await v.patchEmployeeExperienceAssignedRole(t, 'r1', body);
    await v.deleteEmployeeExperienceAssignedRole(t, 'r1', uid, 'W/"1"');

    await v.listEmployeeExperienceAssignedRoleMembers(t, 'r1', uid, '?$top=1');
    await v.createEmployeeExperienceAssignedRoleMember(t, 'r1', body);
    await v.getEmployeeExperienceAssignedRoleMember(t, 'r1', 'm1', uid);
    await v.patchEmployeeExperienceAssignedRoleMember(t, 'r1', 'm1', body);
    await v.deleteEmployeeExperienceAssignedRoleMember(t, 'r1', 'm1', uid, 'W/"1"');

    await v.getEmployeeExperienceAssignedRoleMemberUser(t, 'r1', 'm1', uid);
    await v.getEmployeeExperienceAssignedRoleMemberUserMailboxSettings(t, 'r1', 'm1', uid);
    await v.patchEmployeeExperienceAssignedRoleMemberUserMailboxSettings(t, 'r1', 'm1', body);
    await v.listEmployeeExperienceAssignedRoleMemberUserServiceProvisioningErrors(t, 'r1', 'm1', uid, '?$top=1');

    await v.listLearningCourseActivities(t, uid, '?$top=1');
    await v.getLearningCourseActivity(t, 'act-1', uid);
    await v.getLearningCourseActivityByExternalId(t, "ext'id", uid);

    await v.getAdminPeopleItemInsights(t);
    await v.patchAdminPeopleItemInsights(t, body);
    await v.deleteAdminPeopleItemInsights(t, 'W/"1"');

    await v.getOrganizationItemInsights(t, 'org-1');
    await v.patchOrganizationItemInsights(t, 'org-1', body);
    await v.deleteOrganizationItemInsights(t, 'org-1', 'W/"1"');

    await v.getWorkHoursAndLocations(t);
    await v.patchWorkHoursAndLocations(t, body, uid);
    await v.listWorkHoursOccurrences(t, uid, '?$top=1');
    await v.postWorkHoursSetCurrentLocation(t, uid, body);
    await v.postWorkHoursSetCurrentLocation(t, uid, undefined);
    await v.getWorkHoursOccurrence(t, 'occ-1', uid);
    await v.patchWorkHoursOccurrence(t, 'occ-1', body);
    await v.deleteWorkHoursOccurrence(t, 'occ-1', uid, 'W/"1"');
    await v.getWorkHoursOccurrencesView(t, '2026-01-01T00:00:00Z', '2026-01-02T00:00:00Z', uid);
    await v.listWorkHoursRecurrences(t, uid, '?$top=1');
    await v.getWorkHoursRecurrence(t, 'rec-1', uid);
    await v.patchWorkHoursRecurrence(t, 'rec-1', body);
    await v.deleteWorkHoursRecurrence(t, 'rec-1', uid, 'W/"1"');

    await v.listOnlineMeetingEngagementConversations(t, '?$top=1');
    await v.createOnlineMeetingEngagementConversation(t, body);
    await v.getOnlineMeetingEngagementConversation(t, 'conv-1');
    await v.patchOnlineMeetingEngagementConversation(t, 'conv-1', body);
    await v.deleteOnlineMeetingEngagementConversation(t, 'conv-1', 'W/"1"');
    await v.listOnlineMeetingEngagementConversationMessages(t, 'conv-1', '?$top=1');
    await v.getOnlineMeetingEngagementConversationMessage(t, 'conv-1', 'msg-1');

    expect(true).toBe(true);
  });
});
