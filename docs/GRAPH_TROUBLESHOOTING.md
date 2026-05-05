# Microsoft Graph troubleshooting (m365-agent-cli)

Short notes for **advanced OData** and **people** queries. For delegated scope coverage see [GRAPH_SCOPES.md](./GRAPH_SCOPES.md).

## `ConsistencyLevel: eventual`

Some Graph queries (notably **`$search`**, **`$count=true`** on certain collections, and advanced **`$filter`** combinations) require the request header **`ConsistencyLevel: eventual`**. The CLI sets this where it wraps the API (for example directory **`$count`** and To Do **`--count`** on the single-page path). For **`graph invoke`**, pass headers explicitly, for example:

```bash
m365-agent-cli graph invoke -X GET "/me/messages?\$search=foo" --header "ConsistencyLevel: eventual"
```

Always confirm the current requirement in [Microsoft Graph documentation](https://learn.microsoft.com/en-us/graph/api/overview) for the resource you are calling.

## `/me/people` and `$search`

People relevance search (`GET /me/people` with **`$search=`**) can return **tenant-specific 400s** or empty sets depending on mailbox indexing, policy, and **eventual** consistency. If a call fails:

1. Verify **`People.Read`** (or the documented scope for that API revision) on the token (`verify-token`).
2. Re-read the latest Graph docs for **`/me/people`** and **`$search`** — behavior and header requirements change over time.
3. Prefer higher-level commands (**`find`**, **`contacts`**) when they already model the headers and paths you need.

Do **not** add **`ConsistencyLevel: eventual`** to shared client code for `searchPeople` unless product docs confirm it is broadly required; start with documentation and scoped **`graph invoke`** experiments.

## Path inventory and OpenAPI compliance (optional)

The repo keeps a machine-readable list of Graph-relative paths used by **`callGraph`**, **`graphInvoke`**, **`fetchAllPages`**, **`fetchGraphRaw`**, and **`graphPostBatch`** in [`docs/GRAPH_PATH_INVENTORY.json`](./GRAPH_PATH_INVENTORY.json). Regenerate after changing those call sites: **`npm run graph:inventory`**. CI runs **`npm run graph:inventory:check`** so the file cannot drift silently.

With the [msgraph Cursor skill](https://github.com/merill/msgraph) installed locally (`~/.cursor/skills/msgraph/scripts/run.sh`), run **`npm run verify:graph-openapi:strict`** to cross-check unique patterns against the skill’s OpenAPI index (no live Graph calls). Override the launcher with **`MSGRAPH_SKILL_RUN_SH`**, or set **`GRAPH_OPENAPI_VERIFY=0`** to skip. Patterns that the index cannot match but are still valid are listed in [`scripts/graph-openapi-allowlist.json`](../scripts/graph-openapi-allowlist.json).
