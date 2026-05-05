#!/usr/bin/env node
/**
 * Copy the bundled OpenClaw skill (`skills/m365-agent-cli`) into a target skills directory.
 *
 * Opt-in postinstall: when `npm_lifecycle_event` is `postinstall` and `OPENCLAW_SKILLS_DIR`
 * is unset or empty, this script exits 0 without doing anything.
 *
 * Manual: `node node_modules/m365-agent-cli/scripts/install-openclaw-skill.mjs ~/.openclaw/workspace/skills`
 *     or: `OPENCLAW_SKILLS_DIR=... npm install` (postinstall runs this file).
 */
import { cpSync, existsSync, statSync } from 'node:fs';
import { homedir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = dirname(fileURLToPath(import.meta.url));
const pkgRoot = join(__dirname, '..');
const sourceSkill = join(pkgRoot, 'skills', 'm365-agent-cli');

const isPostinstall = process.env.npm_lifecycle_event === 'postinstall';

/** @param {string} p */
function expandUser(p) {
  if (p === '~') {
    return homedir();
  }
  if (p.startsWith('~/') || p.startsWith('~\\')) {
    return join(homedir(), p.slice(2));
  }
  return p;
}

/**
 * @param {string} skillsRoot Directory that should contain `m365-agent-cli/` (e.g. OpenClaw workspace `skills`)
 */
export function copyOpenclawSkill(skillsRoot) {
  const root = resolve(expandUser(skillsRoot.trim()));
  if (!existsSync(sourceSkill) || !statSync(sourceSkill).isDirectory()) {
    throw new Error(`bundled skill missing at ${sourceSkill}`);
  }
  const dest = join(root, 'm365-agent-cli');
  cpSync(sourceSkill, dest, { recursive: true });
  return dest;
}

function main() {
  const envDir = process.env.OPENCLAW_SKILLS_DIR?.trim();
  const argDir = process.argv[2]?.trim();

  if (isPostinstall && !envDir) {
    return;
  }

  const target = envDir || argDir;
  if (!target) {
    console.error(
      'install-openclaw-skill: pass the OpenClaw skills directory as the first argument, or set OPENCLAW_SKILLS_DIR.\n' +
        'Example: node scripts/install-openclaw-skill.mjs ~/.openclaw/workspace/skills\n' +
        'Example: OPENCLAW_SKILLS_DIR=~/.openclaw/workspace/skills npm install m365-agent-cli'
    );
    process.exit(1);
  }

  try {
    const dest = copyOpenclawSkill(target);
    if (!isPostinstall) {
      console.log(`install-openclaw-skill: copied bundled skill to ${dest}`);
    }
  } catch (e) {
    const msg = /** @type {Error} */ (e).message;
    console.error(`install-openclaw-skill: ${msg}`);
    process.exit(1);
  }
}

const selfPath = resolve(fileURLToPath(import.meta.url));
const invokedPath = process.argv[1] ? resolve(process.argv[1]) : '';
if (invokedPath === selfPath) {
  main();
}
