import { describe, expect, it } from 'bun:test';
import http from 'node:http';
import { startWebhookServer } from './webhook-server.js';

function httpRequest(
  port: number,
  path: string,
  opts?: { method?: string; headers?: Record<string, string>; body?: string }
): Promise<{ status: number; body: string }> {
  return new Promise((resolve, reject) => {
    const req = http.request(
      {
        hostname: '127.0.0.1',
        port,
        path,
        method: opts?.method ?? 'GET',
        headers: opts?.headers
      },
      (res) => {
        const chunks: Buffer[] = [];
        res.on('data', (c) => chunks.push(c as Buffer));
        res.on('end', () => resolve({ status: res.statusCode ?? 0, body: Buffer.concat(chunks).toString('utf8') }));
      }
    );
    req.on('error', reject);
    if (opts?.body) req.write(opts.body);
    req.end();
  });
}

describe('webhook-server', () => {
  it('replays validationToken and returns 404 for unknown paths', async () => {
    const server = startWebhookServer(0);
    const addr = server.address();
    const port = typeof addr === 'object' && addr && 'port' in addr ? addr.port : 0;
    try {
      const v = await httpRequest(port, '/webhooks/m365-agent-cli?validationToken=abc123');
      expect(v.status).toBe(200);
      expect(v.body).toBe('abc123');

      const clippy = await httpRequest(port, '/webhooks/clippy?validationToken=xyz');
      expect(clippy.status).toBe(200);
      expect(clippy.body).toBe('xyz');

      const n404 = await httpRequest(port, '/other');
      expect(n404.status).toBe(404);

      const postFree = await httpRequest(port, '/webhooks/m365-agent-cli', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ note: 'ping' })
      });
      expect(postFree.status).toBe(202);
    } finally {
      await new Promise<void>((resolve) => server.close(() => resolve()));
    }
  });

  it('accepts POST notifications and enforces clientState when configured', async () => {
    const prev = process.env.GRAPH_CLIENT_STATE;
    process.env.GRAPH_CLIENT_STATE = 'expected';

    const server = startWebhookServer(0);
    const addr = server.address();
    const port = typeof addr === 'object' && addr && 'port' in addr ? addr.port : 0;
    try {
      const bad = await httpRequest(port, '/webhooks/m365-agent-cli', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ value: [{ clientState: 'wrong' }] })
      });
      expect(bad.status).toBe(401);

      const good = await httpRequest(port, '/webhooks/m365-agent-cli', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ value: [{ clientState: 'expected' }] })
      });
      expect(good.status).toBe(202);

      const malformed = await httpRequest(port, '/webhooks/m365-agent-cli', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: '{'
      });
      expect(malformed.status).toBe(400);
    } finally {
      await new Promise<void>((resolve) => server.close(() => resolve()));
      if (prev === undefined) delete process.env.GRAPH_CLIENT_STATE;
      else process.env.GRAPH_CLIENT_STATE = prev;
    }
  });
});
