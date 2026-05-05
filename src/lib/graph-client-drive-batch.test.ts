import { describe, expect, it } from 'bun:test';
import { mkdtemp, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import type { DriveLocation } from './drive-location.js';

const token = 'tok';
const baseUrl = 'https://graph.microsoft.com/v1.0';
const item = 'item-gc';
const meItemBase = `${baseUrl}/me/drive/items/${item}`;

describe('graph-client drive item batch', () => {
  const loc: DriveLocation = { kind: 'me' };

  it('covers delete, share, checkout, listItem, follow, sensitivity, restore, retention, versions, delta, sharedWithMe, copy, move, permissions, invite, analytics', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async (input, init) => {
        const url = typeof input === 'string' ? input : input instanceof Request ? input.url : String(input);
        const method = (init?.method || 'GET').toUpperCase();

        if (method === 'GET' && (url === meItemBase || url.startsWith(`${meItemBase}?`))) {
          return new Response(
            JSON.stringify({ id: item, name: 'report.xlsx', webUrl: 'https://example.invalid/doc' }),
            { status: 200, headers: { 'content-type': 'application/json' } }
          );
        }

        if (method === 'GET' && url.includes('/analytics/allTime')) {
          return new Response(JSON.stringify({ access: { actionCount: 1 } }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }
        if (method === 'GET' && url.includes('/analytics/lastSevenDays')) {
          return new Response(JSON.stringify({ access: { actionCount: 2 } }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }

        if (method === 'DELETE' && url === meItemBase) {
          return new Response(null, { status: 204 });
        }

        if (url.includes('/createLink') && method === 'POST') {
          return new Response(JSON.stringify({ link: { webUrl: 'https://share' } }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }

        if (url.includes('/checkout') && method === 'POST') return new Response(null, { status: 204 });
        if (url.includes('/checkin') && method === 'POST') return new Response(null, { status: 204 });
        if (url.includes('/follow') && method === 'POST') return new Response(null, { status: 204 });
        if (url.includes('/unfollow') && method === 'POST') return new Response(null, { status: 204 });
        if (url.includes('/permanentDelete') && method === 'POST') return new Response(null, { status: 204 });
        if (url.includes('/assignSensitivityLabel') && method === 'POST') {
          return new Response(JSON.stringify({ ok: true }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }
        if (url.includes('/extractSensitivityLabels') && method === 'POST') {
          return new Response(JSON.stringify({ labels: [] }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }
        if (url.includes('/restore') && method === 'POST' && url.includes(item)) {
          return new Response(JSON.stringify({ id: item, name: 'restored.xlsx' }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }
        if (url.includes('/retentionLabel') && method === 'GET') {
          return new Response(JSON.stringify({ name: 'label' }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }
        if (url.includes('/retentionLabel') && method === 'DELETE') {
          return new Response(null, { status: 204 });
        }

        if (url.includes('/listItem') && method === 'GET') {
          return new Response(JSON.stringify({ id: 'li1' }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }

        if (url.includes('/versions') && method === 'GET') {
          return new Response(JSON.stringify({ value: [] }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }
        if (url.includes('/versions/v1/restoreVersion') && method === 'POST') {
          return new Response(null, { status: 204 });
        }

        if (url.includes('/copy') && method === 'POST' && url.includes(item)) {
          return new Response(null, {
            status: 202,
            headers: { Location: `${baseUrl}/monitor/copy-job` }
          });
        }

        if (url === `${baseUrl}/me/drive/sharedWithMe` && method === 'GET') {
          return new Response(JSON.stringify({ value: [{ id: 's1' }] }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }

        if (url.startsWith(`${baseUrl}/me/drive/root/delta`) && method === 'GET') {
          return new Response(JSON.stringify({ value: [] }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }

        if (url === meItemBase && method === 'PATCH') {
          return new Response(JSON.stringify({ id: item, name: 'moved.xlsx' }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }

        if (url.includes('/permissions/p1') && method === 'PATCH') {
          return new Response(JSON.stringify({ id: 'p1' }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }
        if (url.includes('/permissions/p1') && method === 'DELETE') {
          return new Response(null, { status: 204 });
        }

        if (url.includes('/invite') && method === 'POST') {
          return new Response(JSON.stringify({ value: [] }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }

        if (url.includes('/special/approot') && method === 'GET') {
          return new Response(JSON.stringify({ id: 'app' }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }

        return new Response(JSON.stringify({ error: { message: `unhandled ${method} ${url}` } }), {
          status: 500,
          headers: { 'content-type': 'application/json' }
        });
      }) as typeof fetch;

      const g = await import('./graph-client.js');

      expect((await g.deleteFile(token, item)).ok).toBe(true);
      expect((await g.shareFile(token, item)).ok).toBe(true);
      expect((await g.checkoutFile(token, item)).ok).toBe(true);
      const cin = await g.checkinFile(token, item, 'done');
      expect(cin.ok).toBe(true);

      expect((await g.getDriveItemListItem(token, item)).ok).toBe(true);
      expect((await g.followDriveItem(token, item)).ok).toBe(true);
      expect((await g.unfollowDriveItem(token, item)).ok).toBe(true);
      expect((await g.assignDriveItemSensitivityLabel(token, item, { labelId: 'x' })).ok).toBe(true);
      expect((await g.extractDriveItemSensitivityLabels(token, item)).ok).toBe(true);
      expect((await g.permanentDeleteDriveItem(token, item)).ok).toBe(true);
      expect((await g.restoreDeletedDriveItem(token, item, {})).ok).toBe(true);
      expect((await g.getDriveItemRetentionLabel(token, item)).ok).toBe(true);
      expect((await g.removeDriveItemRetentionLabel(token, item, 'etag')).ok).toBe(true);

      expect((await g.createOfficeCollaborationLink(token, item, { lock: false })).ok).toBe(true);
      const collabLock = await g.createOfficeCollaborationLink(token, item, { lock: true });
      expect(collabLock.ok).toBe(true);

      expect((await g.listFileVersions(token, item)).ok).toBe(true);
      expect((await g.restoreFileVersion(token, item, 'v1')).ok).toBe(true);

      const fa = await g.getFileAnalytics(token, item);
      expect(fa.ok).toBe(true);
      expect(fa.data?.allTime?.access?.actionCount).toBe(1);

      expect((await g.inviteDriveItem(token, item, { recipients: [], requireSignIn: false })).ok).toBe(true);
      expect((await g.deleteDriveItemPermission(token, item, 'p1')).ok).toBe(true);
      expect((await g.patchDriveItemPermission(token, item, 'p1', { roles: ['read'] })).ok).toBe(true);

      expect((await g.listDriveSharedWithMe(token)).ok).toBe(true);

      expect((await g.getDriveItemDeltaPage(token, { location: loc })).ok).toBe(true);
      expect(
        (await g.getDriveItemDeltaPage(token, { location: loc, nextOrDeltaLink: `${baseUrl}/me/drive/root/delta` })).ok
      ).toBe(true);

      const copy = await g.startCopyDriveItem(token, item, { parentReference: { id: 'parent-1' } });
      expect(copy.ok).toBe(true);
      expect(copy.data?.monitorUrl).toContain('monitor');

      expect((await g.moveDriveItem(token, item, { id: 'folder-2' })).ok).toBe(true);

      const fr = await g.fetchGraphRaw(token, '/me/drive/special/approot', { method: 'GET' }, baseUrl);
      expect(fr.status).toBe(200);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('listFiles and uploadFile simple PUT', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const dir = await mkdtemp(join(tmpdir(), 'm365-up-'));
    const localFile = join(dir, 'hello.bin');
    await writeFile(localFile, Buffer.from('hello'));

    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async (input, init) => {
        const url = typeof input === 'string' ? input : input instanceof Request ? input.url : String(input);
        const method = (init?.method || 'GET').toUpperCase();

        if (url.includes('/root/children') && method === 'GET') {
          return new Response(JSON.stringify({ value: [{ id: 'c1', name: 'a.txt' }] }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }
        if (url.includes(':/content') && method === 'PUT') {
          return new Response(JSON.stringify({ id: 'up1', name: 'hello.bin' }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }

        return new Response(JSON.stringify({ error: { message: `unhandled ${method}` } }), {
          status: 500,
          headers: { 'content-type': 'application/json' }
        });
      }) as typeof fetch;

      const g = await import('./graph-client.js');
      const listed = await g.listFiles(token);
      expect(listed.ok).toBe(true);
      expect(listed.data?.length).toBe(1);

      const up = await g.uploadFile(token, localFile);
      expect(up.ok).toBe(true);
      expect(up.data?.name).toBe('hello.bin');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});
