#!/usr/bin/env node
/**
 * Idempotently inject or update the m365-agent-cli section in TOOLS.md using HTML comment markers.
 * Replaces only the marked region so repeated runs do not append duplicate paragraphs.
 *
 * Usage: node scripts/install-tools-md.mjs <path-to-TOOLS.md>
 *    or: npm run install-tools-md -- <path-to-TOOLS.md>
 */
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

export const TOOLS_MD_BEGIN = '<!-- m365-agent-cli:tools-md begin -->';
export const TOOLS_MD_END = '<!-- m365-agent-cli:tools-md end -->';

const __dirname = dirname(fileURLToPath(import.meta.url));
const pkgRoot = join(__dirname, '..');
const snippetPath = join(pkgRoot, 'packaging', 'tools-md-snippet.md');

export function buildToolsMdBlock() {
  let inner;
  try {
    inner = readFileSync(snippetPath, 'utf8').replace(/\r\n/g, '\n').trimEnd();
  } catch {
    throw new Error(`install-tools-md: could not read snippet at ${snippetPath}`);
  }
  if (!inner) {
    throw new Error('install-tools-md: snippet file is empty');
  }
  return `${TOOLS_MD_BEGIN}\n${inner}\n${TOOLS_MD_END}\n`;
}

/**
 * @param {string} toolsPath absolute or relative path to TOOLS.md
 * @param {{ dryRun?: boolean }} [opts]
 * @returns {{ changed: boolean, content: string }}
 */
export function syncToolsMd(toolsPath, opts = {}) {
  const block = buildToolsMdBlock();
  const markerRegex = new RegExp(
    `${escapeRegex(TOOLS_MD_BEGIN)}\\r?\\n[\\s\\S]*?\\r?\\n${escapeRegex(TOOLS_MD_END)}`,
    'm'
  );

  let before;
  try {
    before = readFileSync(toolsPath, 'utf8').replace(/\r\n/g, '\n');
  } catch (e) {
    if (/** @type {NodeJS.ErrnoException} */ (e).code === 'ENOENT') {
      before = '';
    } else {
      throw e;
    }
  }

  let after;
  if (markerRegex.test(before)) {
    after = before.replace(markerRegex, block.trimEnd());
  } else {
    const trimmed = before.replace(/\s+$/, '');
    const sep = trimmed.length > 0 ? '\n\n' : '';
    after = `${trimmed}${sep}${block}`;
  }

  const changed = after !== before;
  if (changed && !opts.dryRun) {
    mkdirSync(dirname(toolsPath), { recursive: true });
    writeFileSync(toolsPath, after, 'utf8');
  }
  return { changed, content: after };
}

/** @param {string} s */
function escapeRegex(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function main() {
  const toolsPath = process.argv[2];
  if (!toolsPath || toolsPath.startsWith('-')) {
    console.error(
      'Usage: node scripts/install-tools-md.mjs <path-to-TOOLS.md>\n   or: npm run install-tools-md -- <path-to-TOOLS.md>'
    );
    process.exit(1);
  }
  try {
    const { changed } = syncToolsMd(toolsPath);
    console.log(
      changed ? `install-tools-md: updated ${toolsPath}` : `install-tools-md: already up to date (${toolsPath})`
    );
  } catch (e) {
    console.error('install-tools-md:', /** @type {Error} */ (e).message);
    process.exit(1);
  }
}

const selfPath = resolve(fileURLToPath(import.meta.url));
const invokedPath = process.argv[1] ? resolve(process.argv[1]) : '';
if (invokedPath === selfPath) {
  main();
}
