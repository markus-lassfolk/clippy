#!/usr/bin/env node
/**
 * Copy the bundled OpenClaw skill (`skills/m365-agent-cli`) into a target skills directory.
 *
 * Opt-in during installs: when no target directory is given and the process is clearly an
 * `install`/`ci` run (or CI), this script exits 0. Bun often omits `npm_lifecycle_event` for root
 * lifecycles, so we also treat `npm_command=install|ci` as install context.
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

/** True when the package manager is installing and skill copy is optional (opt-in via OPENCLAW_SKILLS_DIR). */
function shouldNoopWhenNoTargetDir() {
  const ev = process.env.npm_lifecycle_event;
  if (ev === 'postinstall' || ev === 'prepare') return true;
  // npm/yarn set npm_command for the top-level install session; Bun often omits npm_lifecycle_event for root lifecycles.
  const cmd = process.env.npm_command;
  if (cmd === 'install' || cmd === 'ci') return true;
  if (process.env.CI) return true;
  return false;
}

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

  const target = envDir || argDir;
  if (!target) {
    if (shouldNoopWhenNoTargetDir()) {
      return;
    }
    console.error(
      'install-openclaw-skill: pass the OpenClaw skills directory as the first argument, or set OPENCLAW_SKILLS_DIR.\n' +
        'Example: node scripts/install-openclaw-skill.mjs ~/.openclaw/workspace/skills\n' +
        'Example: OPENCLAW_SKILLS_DIR=~/.openclaw/workspace/skills npm install m365-agent-cli'
    );
    process.exit(1);
  }

  try {
    const dest = copyOpenclawSkill(target);
    const quietLifecycle =
      process.env.npm_lifecycle_event === 'postinstall' || process.env.npm_lifecycle_event === 'prepare';
    if (!quietLifecycle) {
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
