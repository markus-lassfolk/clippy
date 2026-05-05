#!/usr/bin/env node
/**
 * Line coverage for `src/lib/**` only, excluding `src/lib/ews-client.ts` (SOAP surface; tracked separately).
 * Reads Bun lcov output per-file (SF record). When the same source file appears in multiple records
 * (e.g. `bun test --isolate` merging), DA lines are merged with max hit counts so coverage is not
 * under-counted.
 *
 * Set COVERAGE_MIN_LINES_LIB (default 80) for the gate. While ramping CI, set env below target until tests catch up.
 */
import { readFileSync } from 'node:fs';

const min = Number(process.env.COVERAGE_MIN_LINES_LIB ?? '80');
const lcovPath = process.argv[2] ?? 'coverage/lcov.info';

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

let raw;
try {
  raw = readFileSync(lcovPath, 'utf8');
} catch {
  console.error(`check-coverage-lib: missing or unreadable ${lcovPath} (run bun test --coverage first)`);
  process.exit(1);
}

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

let lf = 0;
let lh = 0;
for (const agg of mergedByFile.values()) {
  const fileLf = agg.lf;
  let fileLh = 0;
  for (let lineNo = 1; lineNo <= fileLf; lineNo += 1) {
    if ((agg.lineHits.get(lineNo) ?? 0) > 0) fileLh += 1;
  }
  lf += fileLf;
  lh += fileLh;
}

const pct = lf === 0 ? 100 : (lh / lf) * 100;
console.log(
  `Lib line coverage (src/lib/**, excluding ews-client.ts): ${pct.toFixed(1)}% (${lh}/${lf} lines), minimum ${min}%`
);

if (pct + 1e-9 < min) {
  console.error(`check-coverage-lib: FAILED — raise lib coverage or lower COVERAGE_MIN_LINES_LIB`);
  process.exit(1);
}
