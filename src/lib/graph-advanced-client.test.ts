import { describe, expect, it } from 'bun:test';
import { graphInvoke, graphInvokeText, graphPostBatch, parseGraphInvokeHeaders } from './graph-advanced-client.js';

describe('parseGraphInvokeHeaders', () => {
  it('parses Name: value with first colon as separator', () => {
    expect(parseGraphInvokeHeaders(['ConsistencyLevel: eventual', 'Prefer: outlook.timezone="UTC"'])).toEqual({
      ConsistencyLevel: 'eventual',
      Prefer: 'outlook.timezone="UTC"'
    });
  });

  it('trims name and value', () => {
    expect(parseGraphInvokeHeaders(['  X-Test :  hello  '])).toEqual({ 'X-Test': 'hello' });
  });

  it('rejects line without colon', () => {
    expect(() => parseGraphInvokeHeaders(['bad'])).toThrow(/Invalid --header/);
  });

  it('rejects empty header name', () => {
    expect(() => parseGraphInvokeHeaders([': only-value'])).toThrow(/empty name/);
  });
});

describe('graphPostBatch', () => {
  it('rejects more than 20 requests without calling fetch', async () => {
    const originalFetch = globalThis.fetch;
    let fetchCalled = false;
    try {
      globalThis.fetch = (() => {
        fetchCalled = true;
        return Promise.resolve(new Response('{}', { status: 200 }));
      }) as unknown as typeof fetch;

      const requests = Array.from({ length: 21 }, (_, i) => ({ id: String(i), method: 'GET', url: '/me' }));
      const r = await graphPostBatch('t', { requests });
      expect(r.ok).toBe(false);
      expect(r.error?.code).toBe('InvalidBatch');
      expect(fetchCalled).toBe(false);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('rejects batch body without requests array', async () => {
    const r = await graphPostBatch('t', {} as { requests: unknown[] });
    expect(r.ok).toBe(false);
    expect(r.error?.message).toMatch(/requests/);
  });

  it('posts $batch and returns JSON', async () => {
    process.env.GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async () =>
        new Response(JSON.stringify({ responses: [] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        })) as typeof fetch;
      const r = await graphPostBatch('tok', { requests: [{ id: '1', method: 'GET', url: '/me' }] });
      expect(r.ok).toBe(true);
      expect((r.data as { responses: unknown[] }).responses).toEqual([]);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});

describe('graphInvoke / graphInvokeText', () => {
  const baseUrl = 'https://graph.microsoft.com/v1.0';

  it('graphInvoke returns error when path is not relative', async () => {
    const r = await graphInvoke('tok', { path: 'me', method: 'GET', pinAccessToken: true });
    expect(r.ok).toBe(false);
    expect(r.error?.message).toMatch(/start with \//);
  });

  it('graphInvoke GET /me succeeds', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async () =>
        new Response(JSON.stringify({ id: 'u1' }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        })) as typeof fetch;
      const r = await graphInvoke('tok', { path: '/me', method: 'GET', pinAccessToken: true });
      expect(r.ok).toBe(true);
      expect((r.data as { id: string }).id).toBe('u1');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('graphInvoke POST sends JSON body', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    let body = '';
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async (_input, init) => {
        body = String(init?.body ?? '');
        return new Response(JSON.stringify({ ok: true }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as typeof fetch;
      const r = await graphInvoke('tok', {
        path: '/me/sendMail',
        method: 'POST',
        body: { x: 1 },
        pinAccessToken: true
      });
      expect(r.ok).toBe(true);
      expect(JSON.parse(body)).toEqual({ x: 1 });
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('graphInvokeText reads non-JSON body', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async () =>
        new Response('plain', {
          status: 200,
          headers: { 'content-type': 'text/plain' }
        })) as typeof fetch;
      const r = await graphInvokeText('tok', { path: '/me', method: 'GET', pinAccessToken: true });
      expect(r.ok).toBe(true);
      expect(r.data).toBe('plain');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});
