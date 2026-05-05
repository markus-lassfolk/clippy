# Agent workflows (m365-agent-cli)

Patterns for **AI agents**, scripts, and orchestrators using Microsoft 365 from the terminal. Prefer **`--json`** (where listed in [CLI_SCRIPTING_INVENTORY.md](./CLI_SCRIPTING_INVENTORY.md)) so stdout is parseable; keep **`stderr`** for human errors.

## 1. Authentication and profiles

- **Config:** `~/.config/m365-agent-cli/` (`.env`, `token-cache-{identity}.json`).
- **Preferred env:** `M365_REFRESH_TOKEN` (or legacy `GRAPH_REFRESH_TOKEN` / `EWS_REFRESH_TOKEN`).
- **Named profile:** `--identity <name>` on commands that support it.
- **One-off token:** `--token <bearer>` (advanced).
- **Sanity checks:** `m365-agent-cli whoami --json`, `m365-agent-cli verify-token --capabilities --json`.

## 2. Read-only safety

- Pass **`--read-only` immediately after** the program name: `m365-agent-cli --read-only files list`.
- Or set **`READ_ONLY_MODE=true`** in the environment or `.env`.
- The authoritative list of blocked actions is in [CLI_REFERENCE.md](./CLI_REFERENCE.md) § Read-Only Mode. Anything not listed may still perform writes if you pass mutating flags—always check subcommand help.

## 3. Drive locations (OneDrive / SharePoint)

For **`files`**, **`excel`**, **`word`**, **`powerpoint`**, use **at most one** of:

- **`--user <upn-or-id>`** — `/users/{id}/drive`
- **`--drive-id`** — explicit drive
- **`--site-id`** — site’s default document library drive
- **`--site-id`** + **`--library-drive-id`** — a specific library on that site

Default is **`/me/drive`**. These flags are **not** the same as EWS **`--mailbox`**.

## 4. Durable sync with delta + `--state-file`

Use a **state file** on disk so the agent can resume incremental sync without re-fetching everything:

| Area | Command | Notes |
| --- | --- | --- |
| To Do | `todo delta --state-file …` | |
| To Do (lists) | `todo lists-delta --state-file …` | Sync **task list** metadata (`lists/delta()`), not tasks in a list. |
| Planner | `planner delta --state-file …` | |
| Contacts | `contacts delta --state-file …` | |
| Mail (Graph) | `outlook-graph messages-delta --state-file …` | |
| Calendar (Graph) | `graph-calendar events-delta --state-file …` | |
| Drive | `files delta --state-file …` | `kind: driveDelta` in state file |

Always persist the updated state file after each successful page.

## 5. Command choice (Graph vs EWS)

| Goal | Prefer | Notes |
| --- | --- | --- |
| Mailbox mail/calendar with **shared mailbox** (`--mailbox`) | `mail`, `calendar`, `send`, … | EWS paths; backend from `M365_EXCHANGE_BACKEND`. |
| **Graph-native** mail ( OData, `$select`, move/copy ) | `outlook-graph` | Distinct from EWS `mail`. |
| **Graph-native** calendar (list view, REST ids) | `graph-calendar` | Invitation responses: `accept` / `decline` / `tentative`. |
| **Files / sharing** | `files` | List, search, upload, share, **invite**, permissions, versions, delta. |
| **Excel** (cells, tables, charts, **workbook comments** beta) | `excel` | |
| **Word / PowerPoint** (preview, metadata, bytes) | `word`, `powerpoint` | See § 7 — no first-class in-file comments like Excel. |
| **Anything else on Graph** | `graph invoke`, `graph batch` | Escape hatch; confirm scopes in [GRAPH_SCOPES.md](./GRAPH_SCOPES.md). |

### 5a. Outlook mail: multiple actions in one script tick

On the **Graph** path, the primary `mail` command applies **one** mutating operation per invocation when using `mail-graph.ts` (combining `--read` with list filters, or stacking `--move` with `--flag` in the same call, is rejected with a hint). For automation that must **PATCH then move** (or similar) without switching to EWS, use either **sequential CLI calls** or a single **`graph batch`** file.

**`graph batch`** — prepare a JSON file `{ "requests": [ { "id": "1", "method": "PATCH", "url": "/me/messages/{id}", "body": { ... }, "headers": { "Content-Type": "application/json" } }, ... ] }` and run:

```bash
m365-agent-cli graph batch -f ./mail-batch.json
```

Keep each `url` **relative to the Graph host** (v1.0 or beta per command flags). Respect the **20 requests per batch** limit. For one-off steps, prefer **`outlook-graph patch-message`** and **`outlook-graph move-message`** instead of batching.

## 6. Teams + files: cross-product collaboration

Graph has no single “notify channel about this file” API. Typical flow:

1. **Resolve the file:** `files meta <itemId> --json` (or `word meta` / search) → read `webUrl` / `name`.
2. **Share or invite:** `files share …`, **`files invite`** with a JSON `--body` file, or **`files permissions`** as needed ([CLI_REFERENCE](./CLI_REFERENCE.md)).
3. **Post to Teams:** `teams channel-message-send … --html '…<a href="WEB_URL">Title</a>…'` or `--text` / `--json-file` for adaptive cards.
4. **@mentions:** use **`--at userObjectId:DisplayName`** (repeatable) with **`--text`** that contains **`@DisplayName`** for each mention (see `teams channel-message-send --help`).

**Limitation:** **`teams chats`** lists **`/me/chats` only**; there is no delegated Graph list of another user’s chats. Use channel messages, or **`graph-search`** with `chatMessage` in **`--types`** / **`--preset extended`** where permitted.

## 7. Word / PowerPoint: edit and review

- **Preview (embed):** `word preview <itemId>` / `powerpoint preview <itemId>` — returns a session URL when the service supports it.
- **Metadata:** `word meta` / `powerpoint meta`.
- **Download bytes:** `word download` / `powerpoint download` (same as `files download` for that item).
- **Thumbnails:** `word thumbnails` / `powerpoint thumbnails` (same as `files thumbnails <itemId>`).
- **All other drive operations** (share, versions, delta, copy, permissions, …): use **`files`** with the same **`--user`** / **`--drive-id`** / **`--site-id`** / **`--library-drive-id`** — see [GRAPH_API_GAPS.md](./GRAPH_API_GAPS.md) Word matrix and [CLI_REFERENCE.md](./CLI_REFERENCE.md).
- **Agent-friendly edit loop:** download → modify locally → **`files upload`** (or large upload) to a folder → optional **`files share`** / **invite** → notify in Teams (§ 6).
- **In-document threaded comments** on Word/PowerPoint are **not** wrapped like **`excel comments-*`**; Graph coverage is limited vs Excel workbooks. See [GRAPH_API_GAPS.md](./GRAPH_API_GAPS.md) and [GRAPH_INVOKE_BOUNDARIES.md](./GRAPH_INVOKE_BOUNDARIES.md) for **beta `graph invoke`** experiments.

## 8. To Do vs Planner vs Outlook

- **To Do:** personal tasks, **`categories[]`**, checklist, linked resources — `todo`.
- **Planner:** plans/buckets, **`category1`–`category6`**, team task boards — `planner`.
- **Outlook:** mail-linked tasks and calendar follow-ups — often `mail` / `calendar` / `outlook-graph`; do not assume category names match between To Do and Outlook master categories (see SKILL).

## 9. Microsoft Search → drive item

1. `m365-agent-cli graph-search '<query>' --json-hits` — stable, flattened hits (`entityType`, `id`, `webUrl`, `name`, …).
2. For **`driveItem`** hits, pass **`id`** into **`files meta`** (with the correct drive flags if not default).

## 10. Cursor / OpenClaw

- Install the bundled skill under **`skills/m365-agent-cli/`** (see repo README / postinstall **`OPENCLAW_SKILLS_DIR`**).
- Optional: **`packages/m365-agent-cli-mcp`** — thin MCP server over stdio that shells to this CLI for a small read-only tool surface ([packages/m365-agent-cli-mcp/README.md](../packages/m365-agent-cli-mcp/README.md)).

## 11. Delegation

For **`--user`** on another person’s mailbox, Teams list, org hierarchy, etc., see [PERSONAL_ASSISTANT_DELEGATION.md](./PERSONAL_ASSISTANT_DELEGATION.md) and [GRAPH_SCOPES.md](./GRAPH_SCOPES.md).
