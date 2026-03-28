import { serve } from 'bun';

export function startWebhookServer(port: number = 3000) {
  console.log(`Starting webhook receiver on http://localhost:${port}/webhooks/clippy`);
  serve({
    port,
    async fetch(req) {
      const url = new URL(req.url);
      if (url.pathname === '/webhooks/clippy') {
        if (req.method === 'GET') {
          const validationToken = url.searchParams.get('validationToken');
          if (validationToken) {
            console.log(`[${new Date().toISOString()}] Received validation token request. Replaying token...`);
            return new Response(validationToken, {
              status: 200,
              headers: { 'Content-Type': 'text/plain' }
            });
          }
          return new Response('Missing validationToken', { status: 400 });
        } else if (req.method === 'POST') {
          try {
            const body = await req.json();
            console.log(`[${new Date().toISOString()}] Received Graph notification:`);
            console.log(JSON.stringify(body, null, 2));
            return new Response('Accepted', { status: 202 });
          } catch (err) {
            console.error('Error parsing notification body:', err);
            return new Response('Bad Request', { status: 400 });
          }
        }
      }
      return new Response('Not Found', { status: 404 });
    }
  });
}
