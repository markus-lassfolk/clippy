#!/usr/bin/env node
import { spawnSync } from 'node:child_process';
/**
 * Optional maintenance: query the local Microsoft Graph OpenAPI index (msgraph Cursor skill)
 * for drive-item + presentation-related paths. If the skill is missing, prints manual steps.
 *
 * @see docs/GRAPH_API_GAPS.md — "PowerPoint Graph API watchlist"
 */
import { existsSync } from 'node:fs';
import { homedir } from 'node:os';
import { join } from 'node:path';

const SKILL_RUN = join(homedir(), '.cursor/skills/msgraph/scripts/run.sh');

function runSearch(query, limit) {
  return spawnSync('bash', [SKILL_RUN, 'openapi-search', '--query', query, '--limit', String(limit)], {
    encoding: 'utf8',
    stdio: ['ignore', 'pipe', 'pipe']
  });
}

if (!existsSync(SKILL_RUN)) {
  console.error(`msgraph skill launcher not found: ${SKILL_RUN}`);
  console.error(
    'Install the msgraph skill, then re-run this script, or run openapi-search manually — see docs/GRAPH_API_GAPS.md (PowerPoint Graph API watchlist).'
  );
  process.exit(0);
}

const queries = [
  ['driveItem presentation', 25],
  ['drives items slide', 15],
  ['items workbook', 5]
];

for (const [q, lim] of queries) {
  console.log(`\n=== openapi-search --query "${q}" --limit ${lim} ===\n`);
  const r = runSearch(q, lim);
  if (r.status !== 0) {
    process.stderr.write(r.stderr || '');
    process.stderr.write(r.stdout || '');
    process.exit(r.status ?? 1);
  }
  console.log((r.stdout || '').trim());
}

console.log(
  '\nInterpretation: many hits for "items workbook" are expected (Excel). New, relevant hits for drive-hosted **presentation** / **slide** APIs on **drive items** may warrant new CLI wrappers — see graph-excel-comments-client.ts as a pattern.\n'
);
