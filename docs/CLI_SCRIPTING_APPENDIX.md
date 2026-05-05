# CLI scripting appendix

This page helps **automation and AI agents** choose machine-readable output and understand which modules participate in **read-only** mode.

## Related docs

- [AGENT_WORKFLOWS.md](./AGENT_WORKFLOWS.md) — auth, deltas, delegation, cross-product recipes (Teams + files).
- [CLI_REFERENCE.md](./CLI_REFERENCE.md) — full command reference and the authoritative **Read-Only Mode** table (which subcommands call `checkReadOnly`).
- [GRAPH_SCOPES.md](./GRAPH_SCOPES.md) — Entra permissions.

## Regenerating the command inventory

The table in [CLI_SCRIPTING_INVENTORY.md](./CLI_SCRIPTING_INVENTORY.md) is **generated** from `src/commands/*.ts` (heuristic for `--json`; `checkReadOnly(` presence). Regenerate after adding commands or flags:

```bash
node scripts/cli-json-readonly-inventory.mjs > docs/CLI_SCRIPTING_INVENTORY.md
```

Root `package.json` exposes **`npm run inventory:scripting`** for the same command.

## Notes on the inventory

- **`graph invoke`** always prints JSON to stdout; the inventory marks **`graph`** as “no `--json`” because there is no separate flag — behavior is JSON-only.
- **`word`** / **`powerpoint`** register subcommands via [src/commands/office-docs-shared.ts](../src/commands/office-docs-shared.ts); the generator treats them as having `--json` when that shared module defines it.
- A **yes** in the `checkReadOnly` column means the file contains at least one guard call; the exact subcommands that are blocked live in **CLI_REFERENCE** (source of truth).
