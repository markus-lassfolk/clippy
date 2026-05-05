# m365-agent-cli-mcp

**Model Context Protocol** (stdio) server that exposes a small, read-biased tool surface on top of globally installed **[m365-agent-cli](https://www.npmjs.com/package/m365-agent-cli)**.

## Prerequisites

- **`m365-agent-cli`** on `PATH` (e.g. `npm install -g m365-agent-cli`) and a completed **`m365-agent-cli login`** for the same user environment the MCP host runs under.
- **Node.js 18+**

Override the binary:

```bash
export M365_AGENT_CLI_BIN=/path/to/bun
# then e.g. exec /path/to/bun run /path/to/m365-agent-cli/src/cli.ts
```

## Install (this package)

From the monorepo:

```bash
cd packages/m365-agent-cli-mcp
npm install
```

## Tools

| Tool | CLI backing |
| --- | --- |
| `m365_whoami` | `whoami --json` |
| `m365_graph_search` | `graph-search <query> --json-hits` (optional `--preset`) |
| `m365_graph_invoke_get` | `--read-only graph invoke -X GET <path>` (optional `beta`) |

Mutating operations are intentionally **not** exposed as separate MCP tools; use the full CLI in a terminal or extend this package locally if your host policy allows it.

**Office on drive (`word`, `powerpoint`, `excel`, `files`):** there are no dedicated MCP tools for these commands. Use **`m365_graph_invoke_get`** for read-only Graph paths, or invoke **`m365-agent-cli`** directly (e.g. `powerpoint preview`, `word upload`, `files permissions`) from a shell the MCP host provides.

## Run

```bash
npx m365-agent-cli-mcp
# or after npm link:
m365-agent-cli-mcp
```

Configure your MCP client (e.g. Cursor) with a **stdio** command pointing at `node /absolute/path/to/packages/m365-agent-cli-mcp/src/server.mjs`.

## See also

- [docs/AGENT_WORKFLOWS.md](../../docs/AGENT_WORKFLOWS.md)
- [docs/CLI_SCRIPTING_APPENDIX.md](../../docs/CLI_SCRIPTING_APPENDIX.md)
