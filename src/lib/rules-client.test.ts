import { describe, expect, it } from 'bun:test';

const token = 'tok';
const baseUrl = 'https://graph.microsoft.com/v1.0';

describe('rules-client', () => {
  it('listMessageRules returns value array', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async () =>
        new Response(
          JSON.stringify({
            value: [{ id: 'r1', displayName: 'R1', priority: 1, isEnabled: true }]
          }),
          { status: 200, headers: { 'content-type': 'application/json' } }
        )) as typeof fetch;

      const { listMessageRules } = await import('./rules-client.js');
      const r = await listMessageRules(token);
      expect(r.ok).toBe(true);
      expect(r.data?.[0]?.id).toBe('r1');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('getMessageRule GETs rule by id', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const urls: string[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        return new Response(JSON.stringify({ id: 'r2', displayName: 'R2', priority: 2, isEnabled: false }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as typeof fetch;

      const { getMessageRule } = await import('./rules-client.js');
      const r = await getMessageRule(token, 'r2', 'delegate@contoso.com');
      expect(r.ok).toBe(true);
      expect(urls[0]).toContain('/users/delegate%40contoso.com/mailFolders/inbox/messageRules/r2');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('createMessageRule POSTs payload', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const bodies: string[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (_input: string | URL | Request, init?: RequestInit) => {
        if (init?.body && typeof init.body === 'string') bodies.push(init.body);
        return new Response(JSON.stringify({ id: 'new', displayName: 'N', priority: 3, isEnabled: true }), {
          status: 201,
          headers: { 'content-type': 'application/json' }
        });
      }) as typeof fetch;

      const { createMessageRule } = await import('./rules-client.js');
      const r = await createMessageRule(token, {
        displayName: 'N',
        actions: { markAsRead: true }
      });
      expect(r.ok).toBe(true);
      expect(JSON.parse(bodies[0] || '{}').displayName).toBe('N');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('updateMessageRule PATCHes and deleteMessageRule DELETEs', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const methods: string[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (_input: string | URL | Request, init?: RequestInit) => {
        methods.push(init?.method || 'GET');
        return new Response(JSON.stringify({ id: 'r', displayName: 'U', priority: 1, isEnabled: true }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as typeof fetch;

      const { updateMessageRule, deleteMessageRule } = await import('./rules-client.js');
      await updateMessageRule(token, 'r', { displayName: 'U' });
      globalThis.fetch = (async () => new Response(null, { status: 204 })) as typeof fetch;
      const d = await deleteMessageRule(token, 'r');
      expect(d.ok).toBe(true);
      expect(methods[0]).toBe('PATCH');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});
