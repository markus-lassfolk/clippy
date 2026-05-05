import { afterEach, describe, expect, test } from 'bun:test';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

/** Keep in sync with `scripts/install-tools-md.mjs` */
const TOOLS_MD_BEGIN = '<!-- m365-agent-cli:tools-md begin -->';
const TOOLS_MD_END = '<!-- m365-agent-cli:tools-md end -->';

const repoRoot = join(dirname(fileURLToPath(import.meta.url)), '../..');
const installToolsMdScript = join(repoRoot, 'scripts/install-tools-md.mjs');
const installOpenclawSkillScript = join(repoRoot, 'scripts/install-openclaw-skill.mjs');

let tmpRoot: string | undefined;

afterEach(() => {
  if (tmpRoot) {
    rmSync(tmpRoot, { recursive: true, force: true });
    tmpRoot = undefined;
  }
});

function runInstallToolsMd(toolsPath: string) {
  return execFileSync(process.execPath, [installToolsMdScript, toolsPath], {
    encoding: 'utf8',
    cwd: repoRoot
  });
}

function runInstallOpenclawSkill(skillsRoot: string) {
  return execFileSync(process.execPath, [installOpenclawSkillScript, skillsRoot], {
    encoding: 'utf8',
    cwd: repoRoot
  });
}

describe('install-tools-md', () => {
  test('inserts a single marked block into a new file and second run is idempotent', () => {
    tmpRoot = mkdtempSync(join(tmpdir(), 'm365-tools-md-'));
    const tools = join(tmpRoot, 'TOOLS.md');

    const out1 = runInstallToolsMd(tools);
    expect(out1).toContain('updated');
    const once = readFileSync(tools, 'utf8');
    expect(once).toContain(TOOLS_MD_BEGIN);
    expect(once).toContain(TOOLS_MD_END);
    expect(once.match(new RegExp(escapeRe(TOOLS_MD_BEGIN), 'g'))?.length).toBe(1);

    const out2 = runInstallToolsMd(tools);
    expect(out2).toContain('already up to date');
    expect(readFileSync(tools, 'utf8')).toBe(once);
  });

  test('replaces inner content between markers without duplicating the block', () => {
    tmpRoot = mkdtempSync(join(tmpdir(), 'm365-tools-md-'));
    const tools = join(tmpRoot, 'TOOLS.md');
    const staleInner = 'OLD CONTENT THAT MUST DISAPPEAR';
    writeFileSync(tools, `# Tools\n\n${TOOLS_MD_BEGIN}\n${staleInner}\n${TOOLS_MD_END}\n\nfooter\n`, 'utf8');

    runInstallToolsMd(tools);
    const body = readFileSync(tools, 'utf8');
    expect(body).not.toContain(staleInner);
    expect(body).toContain('## m365-agent-cli');
    expect(body.match(new RegExp(escapeRe(TOOLS_MD_BEGIN), 'g'))?.length).toBe(1);
  });
});

describe('install-openclaw-skill', () => {
  test('copies bundled skill into target skills root', () => {
    tmpRoot = mkdtempSync(join(tmpdir(), 'm365-skill-'));
    const out = runInstallOpenclawSkill(tmpRoot);
    expect(out).toContain('copied bundled skill');
    const skillMd = join(tmpRoot, 'm365-agent-cli', 'SKILL.md');
    const text = readFileSync(skillMd, 'utf8');
    expect(text.startsWith('---')).toBe(true);
    expect(text).toContain('m365-agent-cli');
  });
});

function escapeRe(s: string) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
