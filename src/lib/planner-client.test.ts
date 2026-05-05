import { describe, expect, it } from 'bun:test';

describe('createPlannerPlanForSignedInUser', () => {
  it('resolves /me then POSTs beta /me/planner/plans with user container', async () => {
    process.env.GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';
    process.env.GRAPH_BETA_URL = 'https://graph.microsoft.com/beta';
    const urls: string[] = [];
    const bodies: string[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        if (init?.body && typeof init.body === 'string') bodies.push(init.body);
        if (urls.length === 1) {
          return new Response(JSON.stringify({ id: 'user-guid-1' }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }
        return new Response(
          JSON.stringify({
            id: 'plan-1',
            title: 'My plan',
            container: {
              url: 'https://graph.microsoft.com/beta/users/user-guid-1',
              type: 'user'
            }
          }),
          { status: 201, headers: { 'content-type': 'application/json' } }
        );
      }) as typeof fetch;

      const { createPlannerPlanForSignedInUser } = await import('./planner-client.js');
      const r = await createPlannerPlanForSignedInUser('tok', 'My plan');

      expect(r.ok).toBe(true);
      expect(r.data?.id).toBe('plan-1');
      expect(urls[0]).toContain('/v1.0/me');
      expect(urls[0]).toContain('$select=id');
      expect(urls[1]).toContain('graph.microsoft.com/beta/me/planner/plans');
      const postBody = bodies[0];
      expect(postBody).toContain('user-guid-1');
      expect(postBody).toContain('"type":"user"');
      expect(postBody).toContain('"title":"My plan"');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});

describe('listPlannerMyDayTasks and listPlannerRecentPlans', () => {
  it('GETs beta /me/planner/myDayTasks and /me/planner/recentPlans when user is omitted', async () => {
    process.env.GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';
    process.env.GRAPH_BETA_URL = 'https://graph.microsoft.com/beta';
    const urls: string[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        return new Response(JSON.stringify({ value: [] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as typeof fetch;

      const { listPlannerMyDayTasks, listPlannerRecentPlans } = await import('./planner-client.js');
      await listPlannerMyDayTasks('tok');
      await listPlannerRecentPlans('tok');

      expect(urls.some((u) => u.includes('/beta/me/planner/myDayTasks'))).toBe(true);
      expect(urls.some((u) => u.includes('/beta/me/planner/recentPlans'))).toBe(true);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('uses /users/{id}/planner/... when user is set', async () => {
    process.env.GRAPH_BETA_URL = 'https://graph.microsoft.com/beta';
    const urls: string[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        return new Response(JSON.stringify({ value: [] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as typeof fetch;

      const { listPlannerMyDayTasks } = await import('./planner-client.js');
      await listPlannerMyDayTasks('tok', 'alice@contoso.com');

      expect(urls[0]).toContain('/beta/users/alice%40contoso.com/planner/myDayTasks');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});

describe('updatePlannerUser', () => {
  it('PATCHes beta …/planner with If-Match and merge body', async () => {
    process.env.GRAPH_BETA_URL = 'https://graph.microsoft.com/beta';
    const calls: { url: string; init?: RequestInit }[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
        calls.push({
          url: typeof input === 'string' ? input : input.toString(),
          init
        });
        return new Response(
          JSON.stringify({ id: 'pu1', '@odata.etag': 'W/"2"', favoritePlanReferences: {} }),
          { status: 200, headers: { 'content-type': 'application/json' } }
        );
      }) as typeof fetch;

      const { updatePlannerUser } = await import('./planner-client.js');
      const r = await updatePlannerUser('tok', undefined, 'W/"1"', { recentPlanReferences: { p1: null } });

      expect(r.ok).toBe(true);
      expect(calls.length).toBe(1);
      expect(calls[0].url).toContain('/beta/me/planner');
      expect(calls[0].init?.method).toBe('PATCH');
      const h = new Headers(calls[0].init?.headers);
      expect(h.get('If-Match')).toBe('W/"1"');
      expect(JSON.parse((calls[0].init?.body as string) || '{}')).toEqual({ recentPlanReferences: { p1: null } });
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});
