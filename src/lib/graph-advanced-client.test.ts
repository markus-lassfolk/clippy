import { describe, expect, it } from 'bun:test';
import { graphPostBatch, parseGraphInvokeHeaders } from './graph-advanced-client.js';

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
});
