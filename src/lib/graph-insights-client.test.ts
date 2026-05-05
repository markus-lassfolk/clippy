import { describe, expect, it } from 'bun:test';

const token = 'tok';
const baseUrl = 'https://graph.microsoft.com/v1.0';

describe('graph-insights-client', () => {
  it('listInsights GETs /me/insights/{kind} with capped $top', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
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

      const { listInsights } = await import('./graph-insights-client.js');
      const r = await listInsights(token, 'trending', { top: 999 });
      expect(r.ok).toBe(true);
      expect(decodeURIComponent(urls[0])).toContain('/me/insights/trending');
      expect(urls[0]).toContain('$top=200');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('listDriveItemActivities GETs activities under drive item', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
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

      const { listDriveItemActivities } = await import('./graph-insights-client.js');
      const r = await listDriveItemActivities(token, { kind: 'me' }, 'item-1', { top: 5 });
      expect(r.ok).toBe(true);
      expect(urls[0]).toContain('/me/drive/items/item-1/activities');
      expect(urls[0]).toContain('$top=5');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('createDriveItemPreview POSTs preview body fields', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const bodies: string[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
        if (init?.body && typeof init.body === 'string') bodies.push(init.body);
        return new Response(JSON.stringify({ getUrl: 'https://preview' }), {
          status: 201,
          headers: { 'content-type': 'application/json' }
        });
      }) as typeof fetch;

      const { createDriveItemPreview } = await import('./graph-insights-client.js');
      const r = await createDriveItemPreview(token, { kind: 'me' }, 'it', {
        page: 1,
        zoom: 2,
        allowEdit: true,
        chromeless: true
      });
      expect(r.ok).toBe(true);
      expect(r.data?.getUrl).toContain('preview');
      const b = JSON.parse(bodies[0] || '{}');
      expect(b.allowEdit).toBe(true);
      expect(b.chromeless).toBe(true);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('listFollowedSites and followSites', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
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

      const { listFollowedSites, followSites } = await import('./graph-insights-client.js');
      await listFollowedSites(token);
      await followSites(token, ['site-a', 'site-b']);
      expect(urls.some((u) => u.includes('/me/followedSites'))).toBe(true);
      expect(urls.some((u) => u.includes('/followedSites/add'))).toBe(true);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('followSites rejects empty ids', async () => {
    const { followSites } = await import('./graph-insights-client.js');
    const r = await followSites(token, []);
    expect(r.ok).toBe(false);
    expect(r.error?.message).toMatch(/at least one site id/i);
  });

  it('unfollowSites returns error when Graph returns per-item failures', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async () =>
        new Response(
          JSON.stringify({
            value: [{ id: 's1', error: { message: 'nope' } }]
          }),
          { status: 207, headers: { 'content-type': 'application/json' } }
        )) as typeof fetch;

      const { unfollowSites } = await import('./graph-insights-client.js');
      const r = await unfollowSites(token, ['s1']);
      expect(r.ok).toBe(false);
      expect(r.error?.message).toContain('Unfollow failed');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});
