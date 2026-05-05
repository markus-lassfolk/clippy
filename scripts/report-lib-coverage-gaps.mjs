#!/usr/bin/env node
/**
 * Print src/lib files (minus ews-client.ts) sorted by uncovered lines (LF−LH).
 * Merges duplicate SF records for the same path (Bun --isolate) like check-coverage-lib.mjs.
 * Usage: node scripts/report-lib-coverage-gaps.mjs [--top N] [coverage/lcov.info]
 */
import { readFileSync } from 'node:fs';

function parseArgs(argv) {
  let top = 40;
  let lcovPath = 'coverage/lcov.info';
  const rest = [];
  for (let i = 2; i < argv.length; i += 1) {
    const a = argv[i];
    if (a === '--top' && argv[i + 1]) {
      top = Math.max(1, Number(argv[i + 1]) || top);
      i += 1;
      continue;
    }
    rest.push(a);
  }
  if (rest.length > 0) lcovPath = rest[0];
  return { top, lcovPath };
}

const { top, lcovPath } = parseArgs(process.argv);

const EWS_EXCLUDE = 'src/lib/ews-client.ts';
const LIB_PREFIX = 'src/lib/';

function normalizeSf(raw) {
  const s = raw.trim().replace(/\\/g, '/');
  if (s.startsWith(LIB_PREFIX)) return s;
  const idx = s.indexOf(LIB_PREFIX);
  if (idx !== -1) return s.slice(idx);
  return s;
}

function includeLibPath(sf) {
  const n = normalizeSf(sf);
  if (!n.startsWith(LIB_PREFIX)) return false;
  if (n === EWS_EXCLUDE) return false;
  if (n.endsWith('.test.ts')) return false;
  return true;
}

const raw = readFileSync(lcovPath, 'utf8');

/** @type {Map<string, { lf: number; lineHits: Map<number, number> }>} */
const mergedByFile = new Map();

for (const block of raw.split('end_of_record')) {
  let sfRaw = '';
  let blockLf = 0;
  const lineHits = new Map();
  for (const line of block.split(/\r?\n/)) {
    const t = line.trim();
    if (t.startsWith('SF:')) sfRaw = t.slice(3).trim();
    if (t.startsWith('LF:')) blockLf = Number(t.slice(3).trim()) || 0;
    if (t.startsWith('DA:')) {
      const rest = t.slice(3);
      const comma = rest.lastIndexOf(',');
      const ln = Number(rest.slice(0, comma));
      const hits = Number(rest.slice(comma + 1));
      if (!Number.isNaN(ln) && !Number.isNaN(hits)) {
        lineHits.set(ln, Math.max(lineHits.get(ln) ?? 0, hits));
      }
    }
  }
  if (!sfRaw || !includeLibPath(sfRaw)) continue;
  const key = normalizeSf(sfRaw);
  if (!mergedByFile.has(key)) {
    mergedByFile.set(key, { lf: 0, lineHits: new Map() });
  }
  const agg = mergedByFile.get(key);
  let maxLine = 0;
  for (const ln of lineHits.keys()) {
    if (ln > maxLine) maxLine = ln;
  }
  agg.lf = Math.max(agg.lf, blockLf, maxLine);
  for (const [ln, hits] of lineHits) {
    agg.lineHits.set(ln, Math.max(agg.lineHits.get(ln) ?? 0, hits));
  }
}

const rows = [];
for (const [file, agg] of mergedByFile) {
  const lf = agg.lf;
  let lh = 0;
  for (let lineNo = 1; lineNo <= lf; lineNo += 1) {
    if ((agg.lineHits.get(lineNo) ?? 0) > 0) lh += 1;
  }
  const unc = lf - lh;
  if (lf > 0) rows.push({ file, lf, lh, unc, pct: ((lh / lf) * 100).toFixed(1) });
}

rows.sort((a, b) => b.unc - a.unc);
console.log(JSON.stringify({ top: rows.slice(0, top), totalFiles: rows.length }, null, 2));
