#!/usr/bin/env node

/**
 * Optional OpenAPI compliance: cross-check docs/GRAPH_PATH_INVENTORY.json path patterns
 * against the msgraph skill OpenAPI index (local FTS). Does not call the live Graph API.
 *
 * Environment:
 *   MSGRAPH_SKILL_RUN_SH   Path to msgraph/scripts/run.sh (default: ~/.cursor/skills/msgraph/scripts/run.sh)
 *   GRAPH_OPENAPI_VERIFY   Set to "0" to skip (exit 0). Default: run when launcher exists.
 *   GRAPH_OPENAPI_STRICT   Set to "1" to exit 1 on any unmatched pattern (after allowlist).
 *
 * Usage:
 *   node scripts/graph-openapi-compliance.mjs [path/to/GRAPH_PATH_INVENTORY.json]
 */

import { spawnSync } from 'node:child_process';
import { existsSync, readFileSync } from 'node:fs';
import { homedir } from 'node:os';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = dirname(fileURLToPath(import.meta.url));
const repoRoot = join(__dirname, '..');

const defaultLauncher = join(homedir(), '.cursor', 'skills', 'msgraph', 'scripts', 'run.sh');
const launcher = process.env.MSGRAPH_SKILL_RUN_SH || defaultLauncher;

const allowlistPath = join(__dirname, 'graph-openapi-allowlist.json');

function loadAllowlist() {
  if (!existsSync(allowlistPath)) return { allowUnmatchedPatterns: [], notes: {} };
  try {
    return JSON.parse(readFileSync(allowlistPath, 'utf8'));
  } catch {
    return { allowUnmatchedPatterns: [], notes: {} };
  }
}

/** Path segments that are only literals (no {dynamic}). */
function staticSegments(pattern) {
  return pattern
    .split('?')[0]
    .split('/')
    .filter((s) => s && !s.includes('{dynamic}'))
    .map((s) => s.replace(/\(\)/g, ''));
}

function longestStaticPrefix(pattern) {
  const noQ = pattern.split('?')[0];
  const i = noQ.indexOf('{dynamic}');
  return i === -1 ? noQ : noQ.slice(0, i);
}

function escapeRegexLiteral(seg) {
  const base = seg.replace(/\(\)$/u, '');
  const hadParen = seg.endsWith('()');
  const esc = base.replace(/[.*+?^${}()|[\]\\]/gu, '\\$&');
  return hadParen ? `${esc}(?:\\(\\))?` : esc;
}

function segmentToRegex(seg) {
  if (seg === '{dynamic}') return '[^/]+';
  return escapeRegexLiteral(seg);
}

/**
 * `{dynamic}/foo` means "any path prefix" + `/foo` (e.g. drive item + `/children`).
 * `/me/x/{dynamic}` means single-segment id only.
 */
function patternToRegex(pattern) {
  let p = pattern.split('?')[0].trim();
  if (!p.startsWith('/')) p = `/${p}`;
  const parts = p.split('/').filter((x) => x !== '');
  if (parts.length === 0) return null;

  if (parts[0] === '{dynamic}') {
    const tailParts = parts.slice(1).map(segmentToRegex);
    if (tailParts.length === 0) return null;
    return new RegExp(`^/.+/${tailParts.join('/')}(?:/)?(?:\\?.*)?$`, 'i');
  }

  const body = parts.map(segmentToRegex).join('/');
  return new RegExp(`^/${body}(?:/)?(?:\\?.*)?$`, 'i');
}

function normalizeCompare(path) {
  return path.split('?')[0].replace(/\(\)/g, '');
}

function pathRoughMatch(inventoryPattern, openapiPath) {
  const pfx = longestStaticPrefix(inventoryPattern);
  if (!pfx || pfx === '/') return false;
  const api = normalizeCompare(openapiPath);
  const pre = normalizeCompare(pfx);
  const pSeg = pre.split('/').filter(Boolean);
  const aSeg = api.split('/').filter(Boolean);
  if (pSeg.length > aSeg.length) return false;
  for (let i = 0; i < pSeg.length; i++) {
    if (pSeg[i] !== aSeg[i]) return false;
  }
  return true;
}

/** me/drive/sharedWithMe aligns with sharedWithMe() under default drive. */
function sharedWithMeMatch(pattern, openapiPath) {
  if (!pattern.includes('sharedWithMe')) return false;
  const api = normalizeCompare(openapiPath);
  return api.includes('sharedWithMe');
}

function searchQueries(pattern) {
  const qs = new Set();
  const base = longestStaticPrefix(pattern).split('?')[0];
  if (base) {
    qs.add(base);
    if (base.startsWith('/')) qs.add(base.slice(1));
  }
  const segs = staticSegments(pattern);
  for (const s of segs) qs.add(s);
  if (segs.length) qs.add(segs.join(' '));
  if (segs.length >= 2) qs.add(segs.slice(-2).join(' '));
  if (segs.length >= 3) qs.add(segs.slice(-3).join(' '));

  if (pattern.includes('/lists') && segs.includes('tasks') && segs.includes('delta')) {
    qs.add('todo lists tasks delta');
  }
  if (pattern === '{dynamic}/lists' || pattern.endsWith('/lists')) {
    qs.add('me todo lists');
  }

  return [...qs].filter(Boolean);
}

const searchCache = new Map();

function openapiSearch(query, limit = 40) {
  const key = `${query}\t${limit}`;
  if (searchCache.has(key)) return searchCache.get(key);
  const r = spawnSync('bash', [launcher, 'openapi-search', '--query', query, '--limit', String(limit)], {
    encoding: 'utf8',
    maxBuffer: 10 * 1024 * 1024
  });
  let parsed;
  try {
    parsed = JSON.parse(r.stdout || '{}');
  } catch {
    parsed = { results: [], _stderr: r.stderr };
  }
  searchCache.set(key, parsed);
  return parsed;
}

function collectPathsForPattern(pattern, maxPaths = 400) {
  const queries = searchQueries(pattern);
  const paths = new Set();
  for (const q of queries) {
    const res = openapiSearch(q, 45);
    const rows = Array.isArray(res.results) ? res.results : [];
    for (const row of rows) {
      if (row.path) paths.add(row.path);
      if (paths.size >= maxPaths) return paths;
    }
  }
  return paths;
}

function verifyPattern(pattern, allowSet) {
  if (allowSet.has(pattern)) {
    return { ok: true, reason: 'allowlist' };
  }
  if (!pattern || pattern === '{dynamic}') {
    return { ok: true, reason: 'skip-dynamic-only' };
  }
  if (pattern.startsWith('http://') || pattern.startsWith('https://')) {
    return { ok: true, reason: 'skip-absolute-url' };
  }

  const re = patternToRegex(pattern);
  const paths = collectPathsForPattern(pattern);

  for (const apiPath of paths) {
    const apiNorm = apiPath.split('?')[0].replace(/\(\)/g, '');
    if (re?.test(apiNorm)) {
      return { ok: true, reason: 'regex-openapi', sample: apiPath };
    }
    if (pathRoughMatch(pattern, apiPath)) {
      return { ok: true, reason: 'prefix-openapi', sample: apiPath };
    }
    if (sharedWithMeMatch(pattern, apiPath)) {
      return { ok: true, reason: 'sharedWithMe-alias', sample: apiPath };
    }
  }

  if (pattern === '/$batch') {
    return { ok: true, reason: 'well-known-json-batch' };
  }

  return { ok: false, reason: 'no-openapi-hit', tried: searchQueries(pattern) };
}

function main() {
  const invPath = join(repoRoot, process.argv[2] || 'docs/GRAPH_PATH_INVENTORY.json');
  if (process.env.GRAPH_OPENAPI_VERIFY === '0') {
    console.log('GRAPH_OPENAPI_VERIFY=0 — skipping OpenAPI compliance.');
    process.exit(0);
  }
  if (!existsSync(launcher)) {
    console.log(`Msgraph skill launcher not found: ${launcher}`);
    console.log('Set MSGRAPH_SKILL_RUN_SH or install the msgraph Cursor skill. Skipping (exit 0).');
    process.exit(0);
  }
  if (!existsSync(invPath)) {
    console.error(`Missing inventory: ${invPath}`);
    process.exit(1);
  }

  const inventory = JSON.parse(readFileSync(invPath, 'utf8'));
  const allow = loadAllowlist();
  const allowSet = new Set(allow.allowUnmatchedPatterns || []);

  const byPattern = new Map();
  for (const e of inventory.entries || []) {
    if (e.kind === 'absolute-url') continue;
    if (!byPattern.has(e.pattern)) byPattern.set(e.pattern, []);
    byPattern.get(e.pattern).push(e);
  }

  const strict = process.env.GRAPH_OPENAPI_STRICT === '1';
  const results = [];

  for (const pattern of [...byPattern.keys()].sort()) {
    const detail = verifyPattern(pattern, allowSet);
    results.push({ pattern, ok: detail.ok, detail });
  }

  const failed = results.filter((r) => !r.ok);
  const okc = results.length - failed.length;

  console.log(
    JSON.stringify(
      {
        launcher,
        inventory: invPath,
        uniquePatternsChecked: results.length,
        matchedOrSkipped: okc,
        failed: failed.length,
        failures: failed.map((f) => ({ pattern: f.pattern, detail: f.detail }))
      },
      null,
      2
    )
  );

  if (failed.length > 0) {
    if (strict) {
      console.error(
        '\nGRAPH_OPENAPI_STRICT=1: add valid patterns to scripts/graph-openapi-allowlist.json or fix paths.'
      );
      process.exit(1);
    }
    console.error(
      `\n${failed.length} pattern(s) had no OpenAPI FTS hit (non-strict: exit 0). Set GRAPH_OPENAPI_STRICT=1 to fail CI.`
    );
  }

  process.exit(0);
}

main();
