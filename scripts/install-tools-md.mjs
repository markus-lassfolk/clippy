#!/usr/bin/env node
/**
 * Idempotent installer for TOOLS.md snippet with replaceable block markers.
 *
 * Usage:
 *   npm run install-tools-md -- <path-to-TOOLS.md>
 *   node scripts/install-tools-md.mjs <path-to-TOOLS.md>
 *
 * Behavior:
 * - If markers exist: Replace content between markers
 * - If markers missing: Append the block at EOF
 * - Creates file if it doesn't exist
 */
import { existsSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = dirname(fileURLToPath(import.meta.url));
const root = join(__dirname, '..');

const BEGIN_MARKER = '<!-- m365-agent-cli:tools-md begin -->';
const END_MARKER = '<!-- m365-agent-cli:tools-md end -->';

function main() {
  const args = process.argv.slice(2);
  if (args.length === 0) {
    console.error('Usage: npm run install-tools-md -- <path-to-TOOLS.md>');
    console.error('   or: node scripts/install-tools-md.mjs <path-to-TOOLS.md>');
    process.exit(1);
  }

  const toolsPath = args[0];
  const snippetPath = join(root, 'dist', 'tools-md-snippet.md');

  if (!existsSync(snippetPath)) {
    console.error(`Error: Snippet file not found: ${snippetPath}`);
    console.error('Make sure the package is properly installed.');
    process.exit(1);
  }

  const snippet = readFileSync(snippetPath, 'utf8');

  // Extract just the content between markers from the snippet
  const snippetMatch = snippet.match(
    new RegExp(`${escapeRegex(BEGIN_MARKER)}([\\s\\S]*?)${escapeRegex(END_MARKER)}`, 'm')
  );
  if (!snippetMatch) {
    console.error('Error: Snippet file is malformed (missing markers)');
    process.exit(1);
  }

  const blockContent = snippetMatch[1];
  const fullBlock = `${BEGIN_MARKER}${blockContent}${END_MARKER}`;

  let content = '';
  let existed = false;

  if (existsSync(toolsPath)) {
    content = readFileSync(toolsPath, 'utf8');
    existed = true;
  }

  const markerRegex = new RegExp(`${escapeRegex(BEGIN_MARKER)}[\\s\\S]*?${escapeRegex(END_MARKER)}`, 'g');

  if (markerRegex.test(content)) {
    // Markers exist - replace the block
    const newContent = content.replace(markerRegex, fullBlock);
    writeFileSync(toolsPath, newContent, 'utf8');
    console.log(`✓ Updated m365-agent-cli section in ${toolsPath}`);
  } else {
    // Markers missing - append at EOF
    const separator = content && !content.endsWith('\n') ? '\n\n' : '\n';
    const newContent = `${content + separator + fullBlock}\n`;
    writeFileSync(toolsPath, newContent, 'utf8');
    if (existed) {
      console.log(`✓ Appended m365-agent-cli section to ${toolsPath}`);
    } else {
      console.log(`✓ Created ${toolsPath} with m365-agent-cli section`);
    }
  }
}

function escapeRegex(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

main();
