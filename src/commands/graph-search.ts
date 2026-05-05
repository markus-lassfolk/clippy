import { Command } from 'commander';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import { flattenMicrosoftSearchHits, microsoftSearchQuery } from '../lib/graph-microsoft-search.js';

const DEFAULT_ENTITY_TYPES = ['message', 'event', 'driveItem', 'listItem', 'person'];

/** Broader verticals — may require extra delegated permissions per entity type (Graph Search docs). */
const EXTENDED_ENTITY_TYPES = [
  ...DEFAULT_ENTITY_TYPES,
  'chatMessage',
  'site',
  'list',
  'acronym',
  'bookmark',
  'externalItem',
  'qna'
];

/** Search verticals that often need Microsoft Search connectors / extra consent — use when you need connectors, QnA, bookmarks, external items. */
const CONNECTORS_ENTITY_TYPES = [
  'externalItem',
  'acronym',
  'qna',
  'bookmark',
  'listItem',
  'list',
  'site',
  'driveItem',
  'message',
  'person'
];

const ENTITY_PRESETS: Record<string, string[]> = {
  default: DEFAULT_ENTITY_TYPES,
  extended: EXTENDED_ENTITY_TYPES,
  connectors: CONNECTORS_ENTITY_TYPES
};

function summarizeResource(r: Record<string, unknown> | undefined): string {
  if (!r) return '(no resource)';
  const type = (r['@odata.type'] as string | undefined)?.replace(/^#microsoft\.graph\./, '') ?? 'item';
  const subject = r.subject as string | undefined;
  const name = r.name as string | undefined;
  const title = r.title as string | undefined;
  const displayName = r.displayName as string | undefined;
  const line = subject || name || title || displayName || (r.id as string) || '';
  return line ? `[${type}] ${line}` : `[${type}]`;
}

export const graphSearchCommand = new Command('graph-search').description(
  'Microsoft Graph Search API (POST /search/query); entity-specific delegated scopes (Mail, Files, Calendars, etc.) — see Graph docs and docs/GRAPH_SCOPES.md'
);

graphSearchCommand
  .argument('<query>', 'Search query string (KQL-style per Graph docs)')
  .option(
    '--preset <name>',
    `Entity bundle: default | extended | connectors (connector-heavy verticals; may need extra permissions)`,
    'default'
  )
  .option('-t, --types <list>', 'Comma-separated entity types (overrides --preset when set)')
  .option('--from <n>', 'Result offset', '0')
  .option('--size <n>', 'Page size (1–1000)', '25')
  .option('--json', 'Output raw JSON response')
  .option('--json-hits', 'Output only flattened hits (stable keys for agents); mutually exclusive with --json')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (
      query: string,
      opts: {
        preset?: string;
        types?: string;
        from?: string;
        size?: string;
        json?: boolean;
        jsonHits?: boolean;
        token?: string;
        identity?: string;
      }
    ) => {
      if (opts.json && opts.jsonHits) {
        console.error('Error: use either --json or --json-hits, not both');
        process.exit(1);
      }
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const parsedTypes = opts.types
        ?.split(',')
        .map((s) => s.trim())
        .filter(Boolean);
      const presetKey = (opts.preset ?? 'default').toLowerCase();
      const presetTypes = ENTITY_PRESETS[presetKey];
      if (!presetTypes && !parsedTypes?.length) {
        console.error(`Error: unknown --preset "${opts.preset}". Use: default, extended, connectors`);
        process.exit(1);
      }
      const entityTypes =
        parsedTypes && parsedTypes.length > 0 ? parsedTypes : (presetTypes ?? [...DEFAULT_ENTITY_TYPES]);
      const from = Math.max(0, parseInt(opts.from ?? '0', 10) || 0);
      const size = Math.min(1000, Math.max(1, parseInt(opts.size ?? '25', 10) || 25));

      const r = await microsoftSearchQuery(auth.token, {
        entityTypes,
        queryString: query,
        from,
        size
      });
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify(r.data, null, 2));
        return;
      }
      if (opts.jsonHits) {
        console.log(JSON.stringify({ hits: flattenMicrosoftSearchHits(r.data) }, null, 2));
        return;
      }

      const blocks = r.data.value ?? [];
      if (blocks.length === 0) {
        console.log('No result blocks (empty value array).');
        return;
      }
      for (const block of blocks) {
        const terms = block.searchTerms?.join(', ') ?? query;
        console.log(`Search terms: ${terms}`);
        const containers = block.hitsContainers ?? [];
        if (containers.length === 0) {
          console.log('  (no hitsContainers)');
          continue;
        }
        for (const c of containers) {
          const hits = c.hits ?? [];
          const total = c.total ?? hits.length;
          console.log(`  Hits (showing ${hits.length}, total reported: ${total})`);
          for (const h of hits) {
            const line = summarizeResource(h.resource);
            const rank = h.rank != null ? `#${h.rank} ` : '';
            console.log(`    ${rank}${line}`);
            if (h.summary?.trim()) {
              const oneLine = h.summary.replace(/\s+/g, ' ').trim().slice(0, 200);
              if (oneLine) console.log(`      ${oneLine}`);
            }
          }
          if (c.moreResultsAvailable)
            console.log('    … more results available (increase --size or paginate with --from)');
        }
      }
    }
  );
