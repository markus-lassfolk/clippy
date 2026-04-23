#!/usr/bin/env node
/**
 * Optional postinstall hook - prints skill path and optionally copies to OPENCLAW_SKILLS_DIR.
 *
 * Behavior:
 * - Always prints the skill path under node_modules
 * - If OPENCLAW_SKILLS_DIR is set: copies skill to that directory (opt-in)
 */
import { copyFileSync, mkdirSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = dirname(fileURLToPath(import.meta.url));
const root = join(__dirname, '..');

function main() {
  const skillPath = join(root, 'skills', 'm365-agent-cli', 'SKILL.md');
  const skillDir = join(root, 'skills', 'm365-agent-cli');
  const readmePath = join(root, 'skills', 'README.md');

  console.log('\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  console.log('📦 m365-agent-cli installed successfully!');
  console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  console.log('\n🤖 OpenClaw Skill available at:');
  console.log(`   ${skillPath}`);
  console.log('\n📖 Skills documentation:');
  console.log(`   ${readmePath}`);

  const targetDir = process.env.OPENCLAW_SKILLS_DIR;

  if (targetDir) {
    console.log('\n🎯 OPENCLAW_SKILLS_DIR detected - installing skill...');
    try {
      const destDir = join(targetDir, 'm365-agent-cli');
      const destSkillPath = join(destDir, 'SKILL.md');

      // Create destination directory
      mkdirSync(destDir, { recursive: true });

      // Copy SKILL.md
      copyFileSync(skillPath, destSkillPath);
      console.log(`   ✓ Copied to: ${destSkillPath}`);
    } catch (err) {
      console.error(`   ✗ Failed to copy skill: ${err.message}`);
      console.error('   You can manually copy using:');
      console.error(`   cp -r ${skillDir} ${targetDir}/`);
    }
  } else {
    console.log('\n💡 To install the skill for OpenClaw:');
    console.log('   mkdir -p ~/.openclaw/workspace/skills');
    console.log(`   cp -r ${skillDir} ~/.openclaw/workspace/skills/`);
    console.log('\n   Or set OPENCLAW_SKILLS_DIR and reinstall:');
    console.log('   export OPENCLAW_SKILLS_DIR=~/.openclaw/workspace/skills');
    console.log('   npm install m365-agent-cli');
  }

  console.log('\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n');
}

main();
