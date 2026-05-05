import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { Command } from 'commander';
import {
  type GraphBatchRequestBody,
  graphInvoke,
  graphPostBatch,
  parseGraphInvokeHeaders
} from '../lib/graph-advanced-client.js';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import { checkReadOnly } from '../lib/utils.js';

function batchHasMutations(body: GraphBatchRequestBody): boolean {
  for (const req of body.requests) {
    const m = String((req as { method?: string }).method || 'GET').toUpperCase();
    if (m !== 'GET' && m !== 'HEAD') return true;
  }
  return false;
}

export const graphCommand = new Command('graph').description(
  'Advanced Microsoft Graph: raw REST invoke and JSON batch ($batch). Paths are relative to GRAPH_BASE_URL (v1.0) or beta.'
);

graphCommand
  .command('invoke')
  .description(
    'Call Graph with a relative path (e.g. /me/messages?$top=5); JSON response only. Advanced OData ($search, some $filter/$count) may need headers such as ConsistencyLevel: eventual — use --header "ConsistencyLevel: eventual" or a higher-level CLI command.'
  )
  .argument('<path>', 'Path starting with / (under v1.0 or beta root)')
  .option('-X, --method <method>', 'HTTP method', 'GET')
  .option('-d, --data <json>', 'JSON request body (for POST/PATCH/PUT)')
  .option('--body-file <path>', 'Read JSON body from file (overrides --data)')
  .option('--beta', 'Use GRAPH_BETA_URL instead of v1.0')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .option(
    '-H, --header <nameValue>',
    'Extra HTTP header ("Name: value", first colon separates name from value). Repeatable, e.g. -H "ConsistencyLevel: eventual"',
    (val: string, prev: string[]) => {
      const acc = prev ?? [];
      acc.push(val);
      return acc;
    },
    [] as string[]
  )
  .action(
    async (
      pathArg: string,
      opts: {
        method: string;
        data?: string;
        bodyFile?: string;
        beta?: boolean;
        token?: string;
        identity?: string;
        header?: string[];
      },
      cmd
    ) => {
      const method = (opts.method || 'GET').toUpperCase();
      if (method !== 'GET' && method !== 'HEAD') {
        checkReadOnly(cmd);
      }

      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }

      let body: unknown | undefined;
      if (opts.bodyFile) {
        const raw = await readFile(resolve(process.cwd(), opts.bodyFile.trim()), 'utf8');
        body = JSON.parse(raw) as unknown;
      } else if (opts.data) {
        body = JSON.parse(opts.data) as unknown;
      }

      let extraHeaders: Record<string, string> | undefined;
      try {
        const lines = opts.header && opts.header.length > 0 ? opts.header : [];
        extraHeaders = lines.length > 0 ? parseGraphInvokeHeaders(lines) : undefined;
      } catch (e) {
        console.error(e instanceof Error ? e.message : String(e));
        process.exit(1);
      }

      const r = await graphInvoke(auth.token, {
        method,
        path: pathArg,
        body,
        beta: opts.beta,
        expectJson: true,
        extraHeaders,
        identity: opts.identity,
        pinAccessToken: !!opts.token
      });

      if (!r.ok) {
        console.error(`Error: ${r.error?.message}`);
        if (r.error?.requestId) {
          console.error(`request-id: ${r.error.requestId}`);
        }
        process.exit(1);
      }
      console.log(JSON.stringify(r.data, null, 2));
    }
  );

graphCommand
  .command('batch')
  .description(
    'POST JSON batch body to /$batch (max 20 requests per call; see https://learn.microsoft.com/en-us/graph/json-batching )'
  )
  .requiredOption('-f, --file <path>', 'JSON file: { "requests": [ { "id", "method", "url", ... }, ... ] }')
  .option('--beta', 'Use GRAPH_BETA_URL')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { file: string; beta?: boolean; token?: string; identity?: string }, cmd) => {
    const raw = await readFile(resolve(process.cwd(), opts.file.trim()), 'utf8');
    const body = JSON.parse(raw) as GraphBatchRequestBody;
    if (batchHasMutations(body)) {
      checkReadOnly(cmd);
    }

    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }

    const r = await graphPostBatch(auth.token, body, opts.beta, {
      identity: opts.identity,
      pinAccessToken: !!opts.token
    });
    if (!r.ok) {
      console.error(`Error: ${r.error?.message}`);
      if (r.error?.requestId) {
        console.error(`request-id: ${r.error.requestId}`);
      }
      process.exit(1);
    }
    console.log(JSON.stringify(r.data, null, 2));
  });
