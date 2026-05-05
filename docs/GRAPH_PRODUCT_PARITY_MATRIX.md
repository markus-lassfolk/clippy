# Graph product parity matrix (m365-agent-cli)

**Purpose:** Map major **Microsoft 365 workloads** to **CLI commands**, Graph coverage (**Implemented** / **Partial** / **Gap**), and notes (delegation, app-only, escape hatch).

**Related:** [`GRAPH_API_GAPS.md`](./GRAPH_API_GAPS.md) (detailed API-by-area gaps + **closure targets** table), [`GRAPH_SCOPES.md`](./GRAPH_SCOPES.md), [`GRAPH_INVOKE_BOUNDARIES.md`](./GRAPH_INVOKE_BOUNDARIES.md) (invoke-only surfaces + recipes), [`CLI_REFERENCE.md`](./CLI_REFERENCE.md) (command inventory), [`AGENT_WORKFLOWS.md`](./AGENT_WORKFLOWS.md) (orchestration patterns), [`CLI_SCRIPTING_APPENDIX.md`](./CLI_SCRIPTING_APPENDIX.md) (`--json` / read-only inventory).

**Legend**

| Status | Meaning |
| --- | --- |
| **Implemented** | Typical CRUD/scripts covered without raw HTTP. |
| **Partial** | Common paths only; long tail via **`graph invoke`** / **`graph batch`**. |
| **Gap** | Not wrapped by design or pending policy review; use **`graph invoke`** or other tooling. |

---

## Workloads

| Workload | CLI surface | Status | Notes |
| --- | --- | --- | --- |
| **Mail / folders / drafts** | `mail`, `folders`, `drafts`, `send`, `outlook-graph` | **Implemented** | Graph + EWS routing via `M365_EXCHANGE_BACKEND`; ids differ between stacks. |
| **Calendar / events** | `calendar` (optional **`--calendar`**), `create-event` (**`--calendar`**), `update-event`, `delete-event`, `graph-calendar` (calendar + group CRUD), `respond`, `findtime` | **Implemented** | Mixed Graph/EWS ids documented in [`GRAPH_EWS_PARITY_MATRIX.md`](./GRAPH_EWS_PARITY_MATRIX.md). |
| **Online meetings** | `meeting`, Teams fields on `create-event` | **Implemented** | |
| **To Do** | `todo` | **Implemented** | **`todo delta`** + **`todo lists-delta`** support **`--state-file`** (tasks vs task **lists**). |
| **Planner** | `planner` | **Implemented** | **`planner delta`** + **`--state-file`**; beta **`plan-archive`** / **`plan-unarchive`**. |
| **Contacts** | `contacts` | **Implemented** | **`contacts delta`** supports **`--state-file`**. |
| **OneNote** | `onenote` | **Implemented** | |
| **Excel (workbook on drive)** | `excel` | **Implemented** | Worksheets, range read/patch/**range-clear**, used-range, **tables** CRUD + rows add/patch/delete + **columns** list/get/patch, **pivot-tables** lifecycle + refresh, **names** list + **name-get** + **worksheet-names** / **worksheet-name-get**, charts, **workbook-get**, **application-calculate**, **session-create** / **session-refresh** / **session-close** (optional **`--session-id`** on mutating calls), **`excel comments-*`** (Graph **beta**); images/shapes/deep range actions → **`graph invoke`**. |
| **Word / PowerPoint (files)** | `word`, `powerpoint` | **Partial** | **`preview`**, **`meta`**, **`download`**, **`thumbnails`** — same drive flags as **`files`**. **Partial** = intentional scope split: deck/body APIs (slide comments, Word in-file comments) are **not in Graph** like Excel **`workbook/comments`**; **`powerpoint`** is **Graph-complete** for what exists (see **`GRAPH_API_GAPS.md`** Word + PowerPoint matrices and **PowerPoint Graph API watchlist**). Lifecycle (**`delta`**, permissions, versions, copy, …) → **`files`**; beta paths → **`graph invoke`**. |
| **OneDrive / SharePoint files** | `files`, `sharepoint`, `site-pages` | **Partial** | **`files`** — list/search/**`delta`**/**`shared-with-me`**/**`thumbnails`**/upload/download/**`copy`**/**`move`**/delete/share/**`invite`**/**`permissions`**/**`permission-remove`**/**`permission-update`**/versions/checkout/convert/analytics + drive flags; **`sharepoint`** **`resolve-site`**, **`get-item`**, **`delete-item`**, **`items-delta`**, **`--json-file`** on create/update; large surface remains **`graph invoke`**. |
| **Teams** | `teams` | **Partial** | **`teams list --user`**; **`channel-files-folder`**; channels/chats; message **send** / **patch** / **delete** (soft + hard + undo-soft); **`activity-notify`**; **`chat-create`**, **`chat-member-add`**; **`team-member-add`**, **`channel-member-add`**; **tabs** list + **tab-get/create/update/delete**; reactions; **`teams chats`** list is **`/me`** only; meeting lifecycle / RSC / admin → **`graph invoke`**. |
| **Presence** | `presence` | **Partial** | Read + set presence for supported delegates. |
| **Bookings** | `bookings` | **Partial** | **staff-availability** app-only per Microsoft. |
| **Microsoft Search** | `graph-search`, `find` | **Partial** | **`--preset`** `default` / `extended` / `connectors` or **`--types`**; exotic verticals may need **`graph invoke`**. |
| **Copilot (Graph)** | `copilot` | **Partial** | Subset of Copilot Graph APIs. |
| **Subscriptions / webhooks** | `subscribe`, `subscriptions`, `serve` | **Implemented** | **`subscriptions renew-all`** for automation. |
| **Directory / people / rooms** | `find`, `org`, `rooms`, `schedule`, `suggest` | **Partial** | **`org manager`** / **`org direct-reports`**; scope/consent dependent. |
| **Discovery / Insights** | `insights`, `files recent`, `files activities`, `files preview`, `sharepoint followed-sites`/`follow`/`unfollow` | **Implemented** | Office Insights (`/me/insights/trending`/`used`/`shared`), drive recent + per-item activity feed, generic drive item preview, SharePoint Following rail. |
| **Outlook Groups** | `groups list`, `conversations`, `thread`, `posts`, `post-reply` | **Implemented** | Microsoft 365 group conversation surface (Outlook Groups inbox); reuses **`Group.ReadWrite.All`**. |
| **Approvals** | `approvals list`, `get`, `steps`, `respond` | **Implemented** | Teams Approvals app + Power Automate approvals via **`/me/approvals`** (beta). New scope **`ApprovalSolution.ReadWrite`** (canonical delegated). |
| **Meeting recordings / transcripts (delegated)** | `meeting recordings`/`recording-download`/`recordings-all`, `meeting transcripts`/`transcript-download`/`transcripts-all` | **Implemented** | Per-meeting + tenant-wide `getAllRecordings` / `getAllTranscripts` (+ delta with `--state-file`). 403 = tenant Stream/Teams policy, not CLI bug. |
| **Teams activity notifications (delegated)** | `teams activity-notify` | **Implemented** | **`POST /me/teamwork/sendActivityNotification`** + **`POST /chats/{id}/sendActivityNotification`**; app-only path stays **`graph invoke`**. |
| **Cloud Communications** (calls, PSTN) | — | **Gap** | **`graph invoke`** only; see [`GRAPH_INVOKE_BOUNDARIES.md`](./GRAPH_INVOKE_BOUNDARIES.md). |

---

## Escape hatch

Any API not wrapped: **`m365-agent-cli graph invoke`** and **`graph batch`** with appropriate delegated scopes in Entra.

---

*Last updated: 2026-05-05 — Excel **`excel`**: pivots, table CRUD, columns, row patch/delete, **range-clear**, **workbook-get**, **application-calculate**, **session-refresh**, names get / worksheet names, **`--session-id`** on mutating calls. Word/PowerPoint: **Graph-complete** narrative + **`scripts/graph-powerpoint-openapi-watch.mjs`**; Phase 1–3 closures: Discovery & Insights, Outlook Groups, Approvals, meeting recordings & transcripts, Teams activity feed; see [`GRAPH_API_GAPS.md`](./GRAPH_API_GAPS.md).*

