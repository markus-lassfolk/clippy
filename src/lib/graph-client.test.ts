import { describe, expect, it } from 'bun:test';
import { writeFileSync } from 'node:fs';
import { mkdtemp, unlink } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { callGraphAt, GraphApiError, uploadLargeFile } from './graph-client.js';

const token = 'test-token';
const baseUrl = 'https://graph.microsoft.com/v1.0';

describe('searchFiles query encoding', () => {
  it('uses driveRootPrefix from DriveLocation for delegated user', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;

    const fetchCalls: string[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request) => {
        fetchCalls.push(typeof input === 'string' ? input : input.toString());
        return new Response(JSON.stringify({ value: [] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as unknown as typeof fetch;

      const { searchFiles } = await import('./graph-client.js');
      await searchFiles(token, 'budget', { kind: 'user', user: 'alice@contoso.com' });

      expect(fetchCalls).toHaveLength(1);
      expect(fetchCalls[0]).toContain("/users/alice%40contoso.com/drive/root/search(q='budget')");
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('encodes single quotes in search query before interpolation', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;

    const fetchCalls: string[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request) => {
        fetchCalls.push(typeof input === 'string' ? input : input.toString());
        return new Response(JSON.stringify({ value: [] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as unknown as typeof fetch;

      const { searchFiles } = await import('./graph-client.js');
      await searchFiles(token, "') and name='anything");

      expect(fetchCalls).toHaveLength(1);
      expect(fetchCalls[0]).toContain("/me/drive/root/search(q='%27%29%20and%20name%3D%27anything')");
      expect(fetchCalls[0]).not.toContain("q=') and name='anything'");
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});

describe('listDriveItemThumbnails', () => {
  it('requests thumbnails under driveItemPath for site library', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;

    const fetchCalls: string[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request) => {
        fetchCalls.push(typeof input === 'string' ? input : input.toString());
        return new Response(JSON.stringify({ value: [{ id: '0', small: { url: 'https://cdn.example/s.png' } }] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as unknown as typeof fetch;

      const { listDriveItemThumbnails } = await import('./graph-client.js');
      const r = await listDriveItemThumbnails(token, 'item-99', {
        kind: 'site',
        siteId: 'contoso.sharepoint.com,a1,b1',
        libraryDriveId: 'lib1'
      });

      expect(r.ok).toBe(true);
      expect(r.data?.length).toBe(1);
      expect(fetchCalls).toHaveLength(1);
      expect(fetchCalls[0]).toContain(
        '/sites/contoso.sharepoint.com%2Ca1%2Cb1/drives/lib1/items/item-99/thumbnails'
      );
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});

describe('uploadLargeFile chunking', () => {
  it('uploads file in chunks and returns DriveItem', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'm365-graph-upload-'));
    const tmpFile = join(dir, 'chunk.bin');
    const fileSize = 25 * 1024 * 1024; // 25MB
    const buffer = new Uint8Array(fileSize);
    buffer.fill(42);
    writeFileSync(tmpFile, buffer);

    const originalFetch = globalThis.fetch;
    const fetchCalls: any[] = [];

    try {
      globalThis.fetch = (async (input: any, init?: any) => {
        const url = typeof input === 'string' ? input : input.toString();

        // 1. Create session POST
        if (url.includes('createUploadSession')) {
          return new Response(
            JSON.stringify({
              uploadUrl: 'https://upload.example.com/session-123',
              expirationDateTime: '2026-04-01T00:00:00.000Z'
            }),
            { status: 200, headers: { 'content-type': 'application/json' } }
          );
        }

        // 2. Chunk PUTs
        if (init?.method === 'PUT') {
          fetchCalls.push({
            url,
            range: (init.headers as any)?.['Content-Range'],
            bodySize: (init.body as any)?.length
          });
          const range = (init.headers as any)?.['Content-Range'];
          if (range?.endsWith('-26214399/26214400')) {
            // Last chunk 10MB*2 to 25MB
            return new Response(JSON.stringify({ id: 'item-123', name: 'test.tmp' }), {
              status: 201,
              headers: { 'content-type': 'application/json' }
            });
          }
          return new Response('{"expirationDateTime": "..."}', { status: 202 });
        }

        return new Response('{}', { status: 200 });
      }) as any;

      const result = await uploadLargeFile('token', tmpFile);

      if (!result.ok) throw new Error(JSON.stringify(result));
      expect(result.data?.driveItem?.id).toBe('item-123');
      expect(fetchCalls.length).toBeGreaterThanOrEqual(3);

      const firstCall = fetchCalls[0];
      expect(firstCall.range).toContain('bytes 0-');
      expect(firstCall.range).toContain('/26214400');

      const lastCall = fetchCalls[fetchCalls.length - 1];
      expect(lastCall.range).toContain('-26214399/26214400');
      expect(lastCall.bodySize).toBeGreaterThan(0);
    } finally {
      globalThis.fetch = originalFetch;
      try {
        await unlink(tmpFile).catch(() => {});
      } catch {}
    }
  });
});

describe('callGraphAt throttling and errors', () => {
  const token = 'test-token';
  const baseUrl = 'https://graph.microsoft.com/v1.0';

  it('retries GET on 429 with Retry-After then succeeds', async () => {
    let n = 0;
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async () => {
        n++;
        if (n === 1) {
          return new Response(JSON.stringify({ error: { code: 'tooManyRequests', message: 'slow' } }), {
            status: 429,
            headers: { 'retry-after': '0', 'content-type': 'application/json' }
          });
        }
        return new Response(JSON.stringify({ ok: true }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as unknown as typeof fetch;

      const r = await callGraphAt<{ ok: boolean }>(baseUrl, token, '/me', { method: 'GET' });
      expect(r.ok).toBe(true);
      expect(r.data?.ok).toBe(true);
      expect(n).toBe(2);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('does not retry POST 429 without Retry-After', async () => {
    let n = 0;
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async () => {
        n++;
        return new Response(JSON.stringify({ error: { code: 'tooManyRequests', message: 'no header' } }), {
          status: 429,
          headers: { 'content-type': 'application/json' }
        });
      }) as unknown as typeof fetch;

      await expect(
        callGraphAt(baseUrl, token, '/me/sendMail', {
          method: 'POST',
          body: JSON.stringify({})
        })
      ).rejects.toBeInstanceOf(GraphApiError);
      expect(n).toBe(1);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('includes request-id in GraphApiError when header present', async () => {
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async () =>
        new Response(JSON.stringify({ error: { message: 'bad', code: 'BadRequest' } }), {
          status: 400,
          headers: {
            'content-type': 'application/json',
            'request-id': 'req-abc-123'
          }
        })) as unknown as typeof fetch;

      try {
        await callGraphAt(baseUrl, token, '/me/x', { method: 'GET' });
        expect.unreachable();
      } catch (e) {
        expect(e).toBeInstanceOf(GraphApiError);
        expect((e as GraphApiError).requestId).toBe('req-abc-123');
      }
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});
