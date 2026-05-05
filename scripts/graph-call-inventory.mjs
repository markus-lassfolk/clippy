#!/usr/bin/env node
/**
 * Static inventory of Microsoft Graph-relative paths used via callGraph*, graphInvoke*,
 * fetchAllPages, fetchGraphRaw, and graphPostBatch.
 *
 * Usage:
 *   node scripts/graph-call-inventory.mjs [--write <file>] [--check <file>]
 *
 * --write   Write JSON inventory (default: docs/GRAPH_PATH_INVENTORY.json)
 * --check   Regenerate and exit 1 if the file differs (CI drift gate)
 */

import { existsSync, readdirSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, join, relative } from 'node:path';
import { fileURLToPath } from 'node:url';
import ts from 'typescript';

const __dirname = dirname(fileURLToPath(import.meta.url));
const repoRoot = join(__dirname, '..');

const TARGET_DIRS = [join(repoRoot, 'src', 'lib'), join(repoRoot, 'src', 'commands')];

const CALLEE_ARG_INDEX = {
  callGraph: 1,
  callGraphAt: 2,
  callGraphAtText: 2,
  callGraphAbsolute: 1,
  fetchAllPages: 1,
  fetchGraphRaw: 1,
  graphPostBatch: -1
};

/** @type {Array<{ file: string, line: number, callee: string, pattern: string, kind: string }>} */
const entries = [];

function walkTsFiles(dir, out = []) {
  for (const e of readdirSync(dir, { withFileTypes: true })) {
    const p = join(dir, e.name);
    if (e.isDirectory()) walkTsFiles(p, out);
    else if (e.isFile() && e.name.endsWith('.ts') && !e.name.includes('.test.ts')) out.push(p);
  }
  return out;
}

function getCalleeName(callExpr) {
  let e = callExpr.expression;
  if (ts.isExpressionWithTypeArguments(e)) e = e.expression;
  if (ts.isIdentifier(e)) return e.text;
  if (ts.isPropertyAccessExpression(e)) return e.name.text;
  return null;
}

/** @returns {{ pattern: string, kind: 'literal' | 'template' | 'dynamic' } | null} */
function extractStringlike(node) {
  if (!node) return null;
  if (ts.isStringLiteral(node) || ts.isNoSubstitutionTemplateLiteral(node)) {
    return { pattern: node.text, kind: 'literal' };
  }
  if (ts.isTemplateExpression(node)) {
    let s = node.head.text;
    for (const span of node.templateSpans) {
      s += '{dynamic}';
      s += span.literal.text;
    }
    return { pattern: s, kind: 'template' };
  }
  return { pattern: '{dynamic}', kind: 'dynamic' };
}

function extractGraphInvokePathArg(obj) {
  if (!ts.isObjectLiteralExpression(obj)) return null;
  for (const prop of obj.properties) {
    if (!ts.isPropertyAssignment(prop)) continue;
    const key = prop.name;
    const name = ts.isIdentifier(key) ? key.text : ts.isStringLiteral(key) ? key.text : null;
    if (name === 'path') return extractStringlike(prop.initializer);
  }
  return null;
}

function record(file, line, callee, pattern, kind) {
  entries.push({
    file: relative(repoRoot, file).replace(/\\/g, '/'),
    line,
    callee,
    pattern,
    kind: callee === 'callGraphAbsolute' ? 'absolute-url' : kind
  });
}

function visitSource(filePath, sourceFile) {
  const visit = (node) => {
    if (ts.isCallExpression(node)) {
      const name = getCalleeName(node);
      if (name === 'graphInvoke' || name === 'graphInvokeText') {
        if (node.arguments.length >= 2) {
          const ex = extractGraphInvokePathArg(node.arguments[1]);
          const line = sourceFile.getLineAndCharacterOfPosition(node.getStart()).line + 1;
          if (ex) record(filePath, line, name, ex.pattern, ex.kind);
          else record(filePath, line, name, '{dynamic}', 'dynamic');
        }
      } else if (name === 'graphPostBatch') {
        const line = sourceFile.getLineAndCharacterOfPosition(node.getStart()).line + 1;
        record(filePath, line, name, '/$batch', 'literal');
      } else if (name && CALLEE_ARG_INDEX[name] !== undefined && CALLEE_ARG_INDEX[name] >= 0) {
        const idx = CALLEE_ARG_INDEX[name];
        if (node.arguments.length > idx) {
          const ex = extractStringlike(node.arguments[idx]);
          const line = sourceFile.getLineAndCharacterOfPosition(node.getStart()).line + 1;
          if (ex) record(filePath, line, name, ex.pattern, ex.kind);
          else record(filePath, line, name, '{dynamic}', 'dynamic');
        }
      }
    }
    ts.forEachChild(node, visit);
  };
  visit(sourceFile);
}

function main() {
  const args = process.argv.slice(2);
  let writePath = join(repoRoot, 'docs', 'GRAPH_PATH_INVENTORY.json');
  let checkPath = null;
  for (let i = 0; i < args.length; i++) {
    if (args[i] === '--write' && args[i + 1]) {
      writePath = join(repoRoot, args[++i]);
    } else if (args[i] === '--check' && args[i + 1]) {
      checkPath = join(repoRoot, args[++i]);
    }
  }

  const files = TARGET_DIRS.filter((d) => existsSync(d)).flatMap((d) => walkTsFiles(d));
  for (const filePath of files) {
    const text = readFileSync(filePath, 'utf8');
    const sf = ts.createSourceFile(filePath, text, ts.ScriptTarget.Latest, true, ts.ScriptKind.TS);
    visitSource(filePath, sf);
  }

  entries.sort((a, b) =>
    a.file !== b.file ? a.file.localeCompare(b.file) : a.line - b.line || a.pattern.localeCompare(b.pattern)
  );

  const payload = {
    version: 1,
    description:
      'Graph path patterns from static analysis of callGraph*, graphInvoke*, fetchAllPages, fetchGraphRaw, graphPostBatch. Regenerate: npm run graph:inventory',
    generatedAt: new Date().toISOString(),
    entryCount: entries.length,
    entries
  };

  const json = `${JSON.stringify(payload, null, 2)}\n`;

  if (checkPath) {
    const existing = readFileSync(checkPath, 'utf8');
    const normalize = (s) => {
      const o = JSON.parse(s);
      delete o.generatedAt;
      return JSON.stringify(o, null, 2);
    };
    if (normalize(existing) !== normalize(json)) {
      console.error(`Graph path inventory drift: ${checkPath} does not match regenerated output.`);
      console.error('Run: npm run graph:inventory');
      process.exit(1);
    }
    console.log(`OK: ${checkPath} matches ${entries.length} entries.`);
    process.exit(0);
  }

  writeFileSync(writePath, json, 'utf8');
  console.log(`Wrote ${entries.length} entries to ${relative(repoRoot, writePath)}`);
}

main();
