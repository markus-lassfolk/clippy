import { readFile } from 'node:fs/promises';
import { Command } from 'commander';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import {
  buildMicrosoftSearchRequest,
  deepMergeSearchRequest,
  flattenMicrosoftSearchHits,
  type MicrosoftSearchQueryResponse,
  microsoftSearchQueryRaw
} from '../lib/graph-microsoft-search.js';

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

function splitCsv(s: string): string[] {
  return s
    .split(',')
    .map((x) => x.trim())
    .filter(Boolean);
}

function collectCliRequestPatch(opts: {
  fields?: string;
  contentSources?: string;
  region?: string;
  aggregationFilters?: string;
  sortJsonFile?: string;
  enableTopResults?: boolean;
  trimDuplicates?: boolean;
}): Record<string, unknown> | undefined {
  const p: Record<string, unknown> = {};
  if (opts.fields?.trim()) p.fields = splitCsv(opts.fields);
  if (opts.contentSources?.trim()) p.contentSources = splitCsv(opts.contentSources);
  if (opts.region?.trim()) p.region = opts.region.trim();
  if (opts.aggregationFilters?.trim()) p.aggregationFilters = splitCsv(opts.aggregationFilters);
  if (opts.enableTopResults) p.enableTopResults = true;
  if (opts.trimDuplicates) p.trimDuplicates = true;
  return Object.keys(p).length ? p : undefined;
}

function assertBodyFileExclusive(
  query: string,
  opts: {
    preset?: string;
    types?: string;
    from?: string;
    size?: string;
    mergeJsonFile?: string;
    fields?: string;
    contentSources?: string;
    region?: string;
    aggregationFilters?: string;
    sortJsonFile?: string;
    enableTopResults?: boolean;
    trimDuplicates?: boolean;
  }
): void {
  if (query.trim()) {
    console.error('Error: omit the query argument when using --body-file');
    process.exit(1);
  }
  const conflicts = [
    opts.mergeJsonFile?.trim(),
    opts.fields?.trim(),
    opts.contentSources?.trim(),
    opts.region?.trim(),
    opts.aggregationFilters?.trim(),
    opts.sortJsonFile?.trim(),
    opts.types?.trim(),
    opts.enableTopResults,
    opts.trimDuplicates,
    (opts.from ?? '0') !== '0',
    (opts.size ?? '25') !== '25',
    (opts.preset ?? 'default').toLowerCase() !== 'default'
  ];
  if (conflicts.some(Boolean)) {
    console.error('Error: --body-file is exclusive; use only --json / --json-hits / --token / --identity with it');
    process.exit(1);
  }
}

function printSearchResponse(
  data: MicrosoftSearchQueryResponse,
  opts: { json?: boolean; jsonHits?: boolean },
  queryLabel: string
): void {
  if (opts.json) {
    console.log(JSON.stringify(data, null, 2));
    return;
  }
  if (opts.jsonHits) {
    console.log(JSON.stringify({ hits: flattenMicrosoftSearchHits(data) }, null, 2));
    return;
  }

  const blocks = data.value ?? [];
  if (blocks.length === 0) {
    console.log('No result blocks (empty value array).');
    return;
  }
  for (const block of blocks) {
    const terms = block.searchTerms?.join(', ') ?? queryLabel;
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
      if (c.moreResultsAvailable) console.log('    … more results available (increase --size or paginate with --from)');
    }
  }
}

export const graphSearchCommand = new Command('graph-search').description(
  'Microsoft Graph Search API (POST /search/query): presets/types, full `searchRequest` shaping (`--merge-json-file`, `--fields`, …), or raw `--body-file` — see Graph docs and docs/GRAPH_SCOPES.md'
);

graphSearchCommand
  .argument('[query]', 'Search query string (KQL-style per Graph docs); omit when using --body-file')
  .option(
    '--preset <name>',
    `Entity bundle: default | extended | connectors (connector-heavy verticals; may need extra permissions)`,
    'default'
  )
  .option('-t, --types <list>', 'Comma-separated entity types (overrides --preset when set)')
  .option('--from <n>', 'Result offset', '0')
  .option('--size <n>', 'Page size (1–1000)', '25')
  .option(
    '--body-file <path>',
    'Full JSON POST body `{ "requests": [ … ] }` for /search/query (exclusive; no query argument)'
  )
  .option(
    '--merge-json-file <path>',
    'Deep-merge JSON into the built searchRequest (after base query/entityTypes/from/size)'
  )
  .option('--fields <list>', 'Comma-separated `fields` (searchRequest)')
  .option('--content-sources <list>', 'Comma-separated `contentSources` (connectors / scoped search)')
  .option('--region <s>', 'Optional `region` hint')
  .option('--aggregation-filters <list>', 'Comma-separated `aggregationFilters`')
  .option('--sort-json-file <path>', 'JSON array for `sortProperty` objects (`name`, `isDescending`)')
  .option('--enable-top-results', 'Set enableTopResults true', false)
  .option('--trim-duplicates', 'Set trimDuplicates true', false)
  .option('--json', 'Output raw JSON response')
  .option('--json-hits', 'Output only flattened hits (stable keys for agents); mutually exclusive with --json')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (
      query: string | undefined,
      opts: {
        preset?: string;
        types?: string;
        from?: string;
        size?: string;
        bodyFile?: string;
        mergeJsonFile?: string;
        fields?: string;
        contentSources?: string;
        region?: string;
        aggregationFilters?: string;
        sortJsonFile?: string;
        enableTopResults?: boolean;
        trimDuplicates?: boolean;
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

      const queryTrim = (query ?? '').trim();
      const bodyFile = opts.bodyFile?.trim();

      if (bodyFile) {
        assertBodyFileExclusive(query ?? '', opts);
        let raw: unknown;
        try {
          raw = JSON.parse(await readFile(bodyFile, 'utf-8'));
        } catch (e) {
          console.error(`Error: could not read --body-file: ${e instanceof Error ? e.message : e}`);
          process.exit(1);
        }
        if (
          !raw ||
          typeof raw !== 'object' ||
          !Array.isArray((raw as { requests?: unknown }).requests) ||
          (raw as { requests: unknown[] }).requests.length === 0
        ) {
          console.error('Error: --body-file must contain JSON with a non-empty "requests" array');
          process.exit(1);
        }
        const r = await microsoftSearchQueryRaw(auth.token, raw as { requests: unknown[] });
        if (!r.ok || !r.data) {
          console.error(`Error: ${r.error?.message}`);
          process.exit(1);
        }
        printSearchResponse(r.data, opts, '(body-file)');
        return;
      }

      if (!queryTrim) {
        console.error('Error: provide a search query or use --body-file');
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

      let requestPatch: Record<string, unknown> | undefined;
      if (opts.mergeJsonFile?.trim()) {
        try {
          const mergeObj = JSON.parse(await readFile(opts.mergeJsonFile.trim(), 'utf-8')) as Record<string, unknown>;
          if (!mergeObj || typeof mergeObj !== 'object' || Array.isArray(mergeObj)) {
            console.error('Error: --merge-json-file must be a JSON object');
            process.exit(1);
          }
          requestPatch = mergeObj;
        } catch (e) {
          console.error(`Error: could not read --merge-json-file: ${e instanceof Error ? e.message : e}`);
          process.exit(1);
        }
      }

      const cliPatch = collectCliRequestPatch(opts);
      if (cliPatch) {
        requestPatch = requestPatch ? deepMergeSearchRequest(requestPatch, cliPatch) : cliPatch;
      }

      if (opts.sortJsonFile?.trim()) {
        let sortProps: unknown;
        try {
          sortProps = JSON.parse(await readFile(opts.sortJsonFile.trim(), 'utf-8'));
        } catch (e) {
          console.error(`Error: could not read --sort-json-file: ${e instanceof Error ? e.message : e}`);
          process.exit(1);
        }
        if (!Array.isArray(sortProps)) {
          console.error('Error: --sort-json-file must contain a JSON array');
          process.exit(1);
        }
        const sortWrap = { sortProperties: sortProps } as Record<string, unknown>;
        requestPatch = requestPatch ? deepMergeSearchRequest(requestPatch, sortWrap) : sortWrap;
      }

      const built = buildMicrosoftSearchRequest({
        entityTypes,
        queryString: queryTrim,
        from,
        size,
        requestPatch
      });

      const r = await microsoftSearchQueryRaw(auth.token, { requests: [built] });
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      printSearchResponse(r.data, opts, queryTrim);
    }
  );
