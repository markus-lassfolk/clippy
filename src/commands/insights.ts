import { Command } from 'commander';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import { type InsightItem, type InsightKind, listInsights } from '../lib/graph-insights-client.js';

export const insightsCommand = new Command('insights').description(
  'Office Graph insights for the signed-in user (delegated): trending, used, shared documents (`/me/insights/...`). Reuses `Sites.ReadWrite.All` / `Files.ReadWrite.All`.'
);

interface InsightsOpts {
  user?: string;
  top?: string;
  json?: boolean;
  token?: string;
  identity?: string;
}

async function runInsights(kind: InsightKind, opts: InsightsOpts): Promise<void> {
  const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
  if (!auth.success || !auth.token) {
    console.error(`Auth error: ${auth.error}`);
    process.exit(1);
  }
  const top = opts.top ? Number.parseInt(opts.top, 10) : undefined;
  if (opts.top && (!Number.isFinite(top) || (top as number) <= 0)) {
    console.error('Error: --top must be a positive integer');
    process.exit(1);
  }
  const r = await listInsights(auth.token, kind, { user: opts.user, top });
  if (!r.ok || !r.data) {
    console.error(`Error: ${r.error?.message ?? 'Insights request failed'}`);
    process.exit(1);
  }
  if (opts.json) {
    console.log(JSON.stringify(r.data, null, 2));
    return;
  }
  const items = r.data.value ?? [];
  if (items.length === 0) {
    console.log(`No insights/${kind} returned.`);
    return;
  }
  for (const it of items) renderInsightLine(kind, it);
}

function renderInsightLine(kind: InsightKind, it: InsightItem): void {
  const v = it.resourceVisualization;
  const ref = it.resourceReference;
  const title = v?.title ?? ref?.id ?? '(untitled)';
  const container = v?.containerDisplayName ? ` — ${v.containerDisplayName}` : '';
  const type = v?.type ? ` [${v.type}]` : '';
  console.log(`${title}${type}${container}`);
  if (ref?.webUrl) console.log(`  ${ref.webUrl}`);
  if (kind === 'used' && it.lastUsed?.lastAccessedDateTime) {
    console.log(`  lastAccessed: ${it.lastUsed.lastAccessedDateTime}`);
  }
  if (kind === 'shared' && it.lastShared) {
    const s = it.lastShared;
    const who = s.sharedBy?.displayName ?? s.sharedBy?.address ?? '';
    if (who) console.log(`  sharedBy: ${who}`);
    if (s.sharedDateTime) console.log(`  sharedAt: ${s.sharedDateTime}`);
    if (s.sharingSubject) console.log(`  subject: ${s.sharingSubject}`);
  }
  if (typeof it.weight === 'number') console.log(`  weight: ${it.weight.toFixed(3)}`);
}

const commonFlags = (cmd: Command) =>
  cmd
    .option('--user <upn-or-id>', 'Target another user (`/users/{id}/insights/...`); requires consent')
    .option('--top <n>', 'Limit results (Graph $top, max 200)')
    .option('--json', 'Output raw Graph JSON')
    .option('--token <token>', 'Use a specific Graph token')
    .option('--identity <name>', 'Graph token cache identity (default: default)');

commonFlags(insightsCommand.command('trending'))
  .description('Documents trending around the user (`GET /me/insights/trending`).')
  .action(async (opts: InsightsOpts) => runInsights('trending', opts));

commonFlags(insightsCommand.command('used'))
  .description('Documents the user used recently (`GET /me/insights/used`).')
  .action(async (opts: InsightsOpts) => runInsights('used', opts));

commonFlags(insightsCommand.command('shared'))
  .description('Documents shared with the user (`GET /me/insights/shared`).')
  .action(async (opts: InsightsOpts) => runInsights('shared', opts));
