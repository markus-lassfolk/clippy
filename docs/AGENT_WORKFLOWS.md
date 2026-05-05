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

### 5b. Personal assistant: weekly planning → approval → calendar commit

For **approval-gated** scheduling (read mail/todos → propose a week → user confirms → write calendar), align orchestration with this CLI’s **actual** verbs:

1. **Read existing calendar** (machine-friendly): `m365-agent-cli calendar nextweek --json` or **`calendar list nextweek --json`** (equivalent; see [CLI_REFERENCE.md](./CLI_REFERENCE.md) § Calendar). For **Graph** `calendarView` / stable REST ids: `m365-agent-cli graph-calendar list-view --start <iso> --end <iso> [--calendar <id>] [--user <upn>]`. For **incremental** sync: `graph-calendar events-delta --state-file …` (§4).
2. **Find gaps / overload:** `findtime`, `schedule`, `suggest` (see [CLI_REFERENCE.md](./CLI_REFERENCE.md)).
3. **After explicit user approval only:** create blocks with **`create-event`** or **`calendar create`** (same flags; one invocation per event is fine), **or** a single **`graph batch -f batch.json`** where each request is `POST` …`/me/events` with a Graph [event](https://learn.microsoft.com/en-us/graph/api/resources/event) JSON body (same **20-request** limit as §5a).
4. **Help:** `m365-agent-cli calendar --help` lists subcommands; list-specific flags are on **`calendar list --help`**.

**Local behavioral tracking** (daily check-ins, CSV/JSON in the PA workspace) does not use this CLI; agents can read that file before step 1 to bias suggestions.

Bundled skill for copy-paste guidance: [skills/m365-agent-cli/SKILL.md](../skills/m365-agent-cli/SKILL.md) § “Weekly planning and approval-gated calendar writes”.

## 6. Teams + files: cross-product collaboration

Graph has no single “notify channel about this file” API. Typical flow:

1. **Resolve the file:** `files meta <itemId> --json` (or `word meta` / search) → read `webUrl` / `name`.
2. **Share or invite:** `files share …`, **`files invite`** with a JSON `--body` file, or **`files permissions`** as needed ([CLI_REFERENCE](./CLI_REFERENCE.md)).
3. **Post to Teams:** `teams channel-message-send … --html '…<a href="WEB_URL">Title</a>…'` or `--text` / `--json-file` for adaptive cards.
4. **@mentions:** use **`--at userObjectId:DisplayName`** (repeatable) with **`--text`** that contains **`@DisplayName`** for each mention (see `teams channel-message-send --help`).

**Limitation:** **`teams chats`** lists **`/me/chats` only**; there is no delegated Graph list of another user’s chats. Use channel messages, or **`graph-search`** with `chatMessage` in **`--types`** / **`--preset extended`** where permitted.

## 7. Word / PowerPoint: edit and review

Graph supports **file lifecycle** and **compliance** on drive items, not a full **in-document** editing API (no Word comment threads or PowerPoint slide/shape REST model like Excel **`workbook/…`**). Use this section to pick a workflow. For a longer guide (checkout, versions, convert, OOXML round-trip), see **[`WORD_POWERPOINT_EDITING.md`](./WORD_POWERPOINT_EDITING.md)**.

### A. Human or browser editing (Office Online)

1. **`word preview` / `powerpoint preview`** or **`word share <id> --collab`** (optionally **`--lock`** with checkout) → open **`webUrl`** / collaboration URL in a browser.
2. **`word checkin <id> --comment "…"`** when using checkout/checkin discipline.

### B. Agent / automation: binary round-trip

1. **`word download` / `powerpoint download`** → edit bytes locally (any tool that writes valid `.docx`/`.pptx`).
2. **`word upload` / `powerpoint upload`** (or **`upload-large`**) back to the same or another folder.
3. Optional **`word convert`** for PDF/HTML export without opening Office.

### C. SharePoint library metadata & compliance

- **`word list-item` / `powerpoint list-item`** — **`GET …/listItem`** (columns, content type); often **404** on personal OneDrive.
- **`word sensitivity-assign`** (`--json-file` per Microsoft Graph) / **`sensitivity-extract`** — Microsoft Purview / MIP; tenant-dependent.
- **`word retention-label`** / **`retention-label-remove`** — retention label on the item.
- **`word follow` / `unfollow`** — pin a file in OneDrive for Business.
- **`word permanent-delete`** — bypass recycle bin where policy allows (destructive).

All of the above exist on **`files`** under the same subcommand names; **`word`/`powerpoint`** are aliases for the same Graph calls.

### D. Discovery & folder context

- **List / delta / search:** **`files list`**, **`files delta`**, **`files search`** (not duplicated under **`word`**).

### E. What Graph does *not* provide here

- **In-document threaded comments** and **slide-level** REST for `.pptx` — see [GRAPH_API_GAPS.md](./GRAPH_API_GAPS.md); **`graph invoke`** only if Microsoft publishes a path for your tenant.
- **Live co-authoring control** — the CLI prepares links and file state; Office Online performs editing.

**Thumbnails:** `word thumbnails` / `powerpoint thumbnails` (same as `files thumbnails <itemId>`).

## 8. To Do vs Planner vs Outlook

- **To Do:** personal tasks, **`categories[]`**, checklist, linked resources — `todo`.
- **Planner:** plans/buckets, **`category1`–`category6`**, team task boards — `planner`.
- **Outlook:** mail-linked tasks and calendar follow-ups — often `mail` / `calendar` / `outlook-graph`; do not assume category names match between To Do and Outlook master categories (see SKILL).

## 9. Microsoft Search → drive item

1. `m365-agent-cli graph-search '<query>' --json-hits` — stable, flattened hits (`entityType`, `id`, `webUrl`, `name`, …).
2. For **`driveItem`** hits, pass **`id`** into **`files meta`** (with the correct drive flags if not default).

## 10. Cursor / OpenClaw

- Install the bundled skill under **`skills/m365-agent-cli/`** (see repo README / postinstall **`OPENCLAW_SKILLS_DIR`**). The skill includes **weekly planning / approval-gated calendar** command vocabulary (§ “Weekly planning…” in that file); workflow steps live in **§5b** above.
- Optional: **`packages/m365-agent-cli-mcp`** — thin MCP server over stdio that shells to this CLI for a small read-only tool surface ([packages/m365-agent-cli-mcp/README.md](../packages/m365-agent-cli-mcp/README.md)).

## 11. Delegation

For **`--user`** on another person’s mailbox, Teams list, org hierarchy, etc., see [PERSONAL_ASSISTANT_DELEGATION.md](./PERSONAL_ASSISTANT_DELEGATION.md) and [GRAPH_SCOPES.md](./GRAPH_SCOPES.md).
