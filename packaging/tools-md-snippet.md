## m365-agent-cli

Bundled **OpenClaw / agent skill** (full command reference): `skills/m365-agent-cli/SKILL.md` under the installed package (e.g. `node_modules/m365-agent-cli/skills/m365-agent-cli/SKILL.md`).

- **CLI on PATH:** install globally (`npm install -g m365-agent-cli`) or use your package manager’s local `node_modules/.bin`. The published entry uses a **Bun** shebang; on machines **without Bun**, run the TypeScript entry with **tsx**, for example: `npx tsx node_modules/m365-agent-cli/src/cli.ts -- --help` (adjust `node_modules/...` for global installs, e.g. `$(npm root -g)/m365-agent-cli/src/cli.ts`).
- **Copy the skill into your OpenClaw workspace** (one-time per machine/version), or set **`OPENCLAW_SKILLS_DIR`** to your skills root and run **`npm install m365-agent-cli`** so **postinstall** copies `m365-agent-cli` into that directory (opt-in; no writes when the variable is unset).
- **Refresh this section without duplicates:** `npm run install-tools-md -- path/to/TOOLS.md` from a git clone, or `node node_modules/m365-agent-cli/scripts/install-tools-md.mjs path/to/TOOLS.md` when installed from npm.

Repository: [m365-agent-cli](https://github.com/markus-lassfolk/m365-agent-cli).
