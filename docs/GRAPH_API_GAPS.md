# Microsoft Graph vs this CLI — capabilities and gaps

**Purpose:** Track **Graph API areas** that are **implemented** in `m365-agent-cli`, **partially** covered, or **not** exposed, so we can prioritize work and set expectations.

**Related:** [`MIGRATION_TRACKING.md`](./MIGRATION_TRACKING.md), [`GRAPH_EWS_PARITY_MATRIX.md`](./GRAPH_EWS_PARITY_MATRIX.md), [`GRAPH_SCOPES.md`](./GRAPH_SCOPES.md), [`GRAPH_PRODUCT_PARITY_MATRIX.md`](./GRAPH_PRODUCT_PARITY_MATRIX.md) (workloads ↔ commands), [`GRAPH_INVOKE_BOUNDARIES.md`](./GRAPH_INVOKE_BOUNDARIES.md) (invoke-only surfaces), [`GRAPH_WRAPPER_GAP_AUDIT.md`](./GRAPH_WRAPPER_GAP_AUDIT.md) (consolidated gap backlog + hardening checklist).

---

## Legend

| Status | Meaning |
| --- | --- |
| **Implemented** | Primary or parallel command covers typical use. |
| **Partial** | Some APIs or flags; not exhaustive vs Graph. |
| **Gap** | Graph supports it; this CLI does not wrap it (use Graph directly, another tool, or contribute). |

---

## Closure targets (parity roadmap)

Measurable exit criteria for workloads still **Partial** / **Gap**. Command references point to [`docs/CLI_REFERENCE.md`](./CLI_REFERENCE.md); scopes to [`GRAPH_SCOPES.md`](./GRAPH_SCOPES.md). For **AI agents** (deltas, `--json`, Teams + files, Word/PPT round-trip), see [`docs/AGENT_WORKFLOWS.md`](./AGENT_WORKFLOWS.md) and the generated [`docs/CLI_SCRIPTING_INVENTORY.md`](./CLI_SCRIPTING_INVENTORY.md).

| Workload | Delegated target verbs / docs | App-only or invoke-only | Exit criteria (Partial → Implemented or documented closure) |
| --- | --- | --- | --- |
| **Excel** | **`excel`** — worksheets, **range** / **range-patch** / **range-clear**, **used-range**, **tables** (list/get + **table-add|patch|delete**, rows add/**table-row-patch|delete**, **table-columns|column-get|column-patch**), **pivot-tables** + **pivot-table-*** refresh, **names** + **name-get** + **worksheet-names** / **worksheet-name-get**, charts, **workbook-get**, **application-calculate**, **session-create|refresh|close**, optional **`--session-id`** on mutating calls, **`comments-*`** (beta) | — | **Implemented** for script/agent workbook automation; workbook **images** / **shapes** / deep **range()** method graph → **`graph invoke`** (see [`CLI_REFERENCE.md`](./CLI_REFERENCE.md)). |
| **Word (.docx)** | **`word`**: full **`files`** item mirror incl. **`list-item`**, **`follow`**/**`unfollow`**, **`sensitivity-assign`**/**`sensitivity-extract`**, **`retention-label`**/**`retention-label-remove`**, **`permanent-delete`** (+ **`preview`** … **`activities`**) | — | **Implemented** for Graph-documented drive-item APIs; **Gap** = in-document threaded comments (no OpenAPI path). Folder **list**/**`delta`**/**`search`** → **`files`**. |
| **PowerPoint (.pptx)** | **`powerpoint`**: same mirror as **`word`** | — | **Implemented** on Graph; **Gap** = slide/body comments / `…/presentation/…` OM; beta → **`graph invoke`** — **[`GRAPH_INVOKE_BOUNDARIES.md`](./GRAPH_INVOKE_BOUNDARIES.md)** + watchlist. |
| **Teams** | **`teams`** — as below + **app-catalog** / **app-catalog-get**, **apps** / **app-get** / **app-add** / **app-patch** / **app-upgrade** / **app-delete**, **chat-apps** / **chat-app-***, **user-apps** / **user-app-*** (personal scope) | RSC / tenant admin | **`teams list --user`** = **`GET /users/{id}/joinedTeams`**; chat **list** stays **`/me/chats`**; **`POST /users/{id}/teamwork/sendActivityNotification`** → **`teams activity-notify --user-id`** (typically app **`--token`**); admin/RSC/shifts → **[`GRAPH_INVOKE_BOUNDARIES.md`](./GRAPH_INVOKE_BOUNDARIES.md)**. |
| **Presence** | **`presence`** — **`me`** / **`user`** / **`bulk`**, session **set/clear**, **`status-message-set`**, **`preferred-set`** / **`preferred-clear`**, **`clear-location`** | — | **subscription** automation via **`subscribe`** + **`serve`** (see command help). |
| **Search** | **`graph-search`**: presets / `--types`, **`searchRequest`** flags + **`--merge-json-file`**, raw **`--body-file`** | — | **Implemented** for `POST /search/query` automation; uncommon **`resultTemplateOptions`** / **`sharePointOneDriveOptions`** shapes → **`--merge-json-file`** or **`--body-file`**. |
| **Files / SharePoint** | **`files`** (delta, **`shared-with-me`**, **`thumbnails`**, copy/move, drive flags, **`invite`**, **`permissions`**, **`permission-remove`**, **`permission-update`**, **`checkout`** / **`checkin`**), **`sharepoint`** (**`resolve-site`**, **`get-site`**, **`drives`**, **`lists`**, **`get-list`**, **`columns`**, **`items`** with OData paging, get/delete item, items-delta, **`--json-file`** on create/update, followed sites), **`site-pages`**, **`excel`** (workbook object model incl. comments beta), **`teams channel-files-folder`** | — | Core drive + site libraries + list schema + list sync + sharing + Excel on-item workbook APIs wrapped; long tail stays **`graph invoke`**. |
| **Copilot** | **`copilot`** | — | Graph `/copilot` OpenAPI-aligned (incl. `$count`, `/copilot/users` **`ai-user`**, root/communications/reports nav, settings/admin deletes, meeting insight mutations, activity-feed root + counts); preview/beta labeled in help. |
| **Directory / rooms** | **`find`**, **`people`**, **`org`**, **`rooms`**, **`schedule`**, **`suggest`** | Admin directory | **Implemented** for delegated **Places** (lists, rooms in list, find w/ **query**, **get** place, **`find --rooms`**) + **People** (**`people list|get`**, **`/users/{id}/people`** with **`--user`**) + **org** (**`user`**, **`transitive-reports`**, manager, direct-reports). Tenant admin directory CRUD → **invoke**. |
| **Bookings** | Full CRUD + reads + **publish** / **unpublish** / business create-delete | **`bookings staff-availability`** | **Implemented** for delegated Bookings v1; **`staff-availability`** stays **app-only** (`--token`). |
| **Cloud Communications** (excl. wrapped meeting recordings/transcripts) | — | **invoke / out of scope** | **PSTN / Teams Phone** is **out of scope** for this CLI. Other communications APIs → **`graph invoke`** if approved — [`GRAPH_INVOKE_BOUNDARIES.md`](./GRAPH_INVOKE_BOUNDARIES.md). |
| **Mail migration** | **`oof`**, **`rules`**, **`update-event`** (Graph id) | **`auto-reply`** (EWS) | **auto-reply** 1:1 Graph template UX = **not** a goal; see [`MIGRATION_TRACKING.md`](./MIGRATION_TRACKING.md). |
| **Discovery / Insights** (Phase 1) | **`insights trending`** / **`used`** / **`shared`**, **`files recent`** / **`files activities`** / **`files preview`**, **`sharepoint followed-sites`** / **`follow`** / **`unfollow`** | — | Wraps **`/me/insights/*`**, **`/me/drive/recent`**, **`/drives/{id}/items/{id}/activities`**, **`/me/drive/items/{id}/preview`**, **`/me/followedSites`** (+`add`/`remove`). Reuses **`Sites.ReadWrite.All`** + **`Files.ReadWrite.All`**. |
| **Outlook Groups** (Phase 2) | **`groups list`** / **`conversations`** / **`thread`** / **`posts`** / **`post-reply`** | — | Wraps **`/me/memberOf`** filtered to Microsoft 365 groups + **`/groups/{id}/conversations/threads/posts/reply`**. Reuses **`Group.ReadWrite.All`**. |
| **Approvals** (Phase 2) | **`approvals list`** / **`get`** / **`steps`** / **`respond`** | — | Wraps **`/me/approvals`** + **`/steps`** (beta). Adds delegated **`ApprovalSolution.ReadWrite`** scope (canonical name; identifier `6768d3af-4562-48ff-82d2-c5e19eb21b9c`). A narrower **`ApprovalSolutionResponse.ReadWrite`** is available for read-and-respond only. |
| **Meeting recordings & transcripts** (Phase 3) | **`meeting recordings`** / **`recording-download`** / **`recordings-all`** (+ `--delta`); **`meeting transcripts`** / **`transcript-download`** / **`transcripts-all`** | — | Wraps **`/me/onlineMeetings/{id}/recordings`** + **`/transcripts`**, **`getAllRecordings(...)`** / **`getAllTranscripts(...)`**, and **`recordings/delta()`** / **`transcripts/delta()`**. Adds **`OnlineMeetingRecording.Read.All`** (transcripts already have **`OnlineMeetingTranscript.Read.All`**). 403 typically means tenant Stream/Teams policy. |
| **Teams activity feed** (Phase 3) | **`teams activity-notify`** | — | Wraps **`POST /me/teamwork/sendActivityNotification`**, **`POST /chats/{id}/sendActivityNotification`**, and **`POST /users/{id}/teamwork/sendActivityNotification`** (**`--user-id`**, typically application **`--token`**). Adds **`TeamsActivity.Send`**. |

---

## Exchange / Outlook (mail & calendar)

| Graph area | CLI | Notes |
| --- | --- | --- |
| Messages CRUD, send, attachments | **Implemented** | `mail`, `send`, `drafts`, `folders`, `outlook-graph`; **large file** attachments use **createUploadSession** + chunked PUT in `addFileAttachmentToMailMessage`; **`send`** uses **draft + send** when an attachment exceeds the threshold |
| Message search / list filters | **Implemented** | `mail` Graph path (`mail-graph.ts`) — not every flag combo in **one** invocation; combine via sequential calls or **`graph batch`** (see [AGENT_WORKFLOWS.md](./AGENT_WORKFLOWS.md) §5a) |
| **messages/delta** (sync) | **Implemented** | **`outlook-graph messages-delta`** — first page or `--next` for `@odata.nextLink` |
| Calendar view, events CRUD | **Implemented** | `calendar`, `create-event`, `update-event`, `delete-event`, `graph-calendar`; **non-default calendar:** `calendar` / `create-event` **`--calendar <id>`**; **`graph-calendar`** **create-calendar** / **update-calendar** / **delete-calendar**, **list-calendar-groups** / **create-calendar-group** / **delete-calendar-group** |
| **events/delta** | **Implemented** | **`graph-calendar events-delta`** — optional `--calendar`; `--next` for paging |
| Attachments on events | **Implemented** | `calendar --list-attachments` / `--download-attachments`; large files use **upload session** inside `addFileAttachmentToCalendarEvent` (same threshold as mail) |
| Calendar sharing (calendarPermission) | **Implemented** | `delegates list`; **`delegates calendar-share add, update, remove`** (Graph model) |
| Classic EWS delegates (folder matrix) | **Implemented (EWS)** | `delegates add, update, remove` — **not** Graph 1:1 |
| Inbox rules | **Implemented** | `rules` |
| Automatic replies (mailboxSettings) | **Implemented** | `oof` |
| **mailboxSettings** (time zone, working hours, formats) | **Implemented** | **`mailbox-settings`** (read + **`set`** with **`--timezone`**, **`--work-*`**, **`--json-file`**) |
| EWS-style auto-reply templates | **Partial** | `auto-reply` (EWS); **`oof`** / **`rules`** for Graph-native |

---

## Word (.docx) drive-hosted Graph — coverage matrix

**Definition of done (Graph “feature complete” for Word here):** every **stable delegated** drive-item operation that applies to `.docx` is **Implemented** via **`files`** and/or **`word`** ( **`word`** mirrors the common per-item **`files`** verbs listed below), or explicitly **Gap / invoke**. Folder-level ops (**list**, **delta**, **search**, **shared-with-me**) stay on **`files`** only.

| Graph area (drive item) | `files` | `word` | `graph invoke` / other |
| --- | --- | --- | --- |
| List / search / delta / shared-with-me | **Implemented** | — | — |
| Metadata (`GET …/items/{id}`) | **Implemented** | **`word meta`** | — |
| Download content | **Implemented** | **`word download`** | — |
| Thumbnails (`GET …/items/{id}/thumbnails`) | **`files thumbnails`** | **`word thumbnails`** (same impl as **`powerpoint thumbnails`**) | — |
| Upload / large upload / delete | **Implemented** | **`word upload`**, **`word upload-large`**, **`word delete`** | — |
| Copy / move | **Implemented** | **`word copy`**, **`word move`** | — |
| Share link / collab / checkout / checkin | **Implemented** | **`word share`**, **`word checkout`**, **`word checkin`** | — |
| Invite / permissions / PATCH permission | **Implemented** | **`word invite`**, **`word permissions`**, **`word permission-remove`**, **`word permission-update`** | — |
| Versions / restore | **Implemented** | **`word versions`**, **`word restore`** | — |
| Convert format | **Implemented** | **`word convert`** | — |
| Analytics | **Implemented** | **`word analytics`** | — |
| Per-item activity feed | **Implemented** | **`word activities`** | — |
| Preview session (`POST …/preview`) | — | **`word preview`** (also **`powerpoint preview`**) | — |
| In-document comments / review (Word-specific) | — | — | **Gap** — no **`…/workbook/comments`–style** Word path in OpenAPI index; Office client / OOXML / beta docs only |
| Sensitivity / MIP labels (assign, extract) | **Implemented** | **`word sensitivity-assign`**, **`sensitivity-extract`** (same as **`files`**) | Tenant licensing + Purview; confirm JSON body in [assignSensitivityLabel](https://learn.microsoft.com/en-us/graph/api/driveitem-assignsensitivitylabel) |
| Retention label (get / remove) | **Implemented** | **`word retention-label`**, **`retention-label-remove`** | [getRetentionLabel](https://learn.microsoft.com/en-us/graph/api/driveitem-getretentionlabel) / [remove](https://learn.microsoft.com/en-us/graph/api/driveitem-removeretentionlabel) |
| SharePoint **listItem** facet | **Implemented** | **`word list-item`** | **`GET …/listItem`** — library columns; often 404 on personal OneDrive |
| Follow file (OneDrive for Business) | **Implemented** | **`word follow`**, **`word unfollow`** | [follow](https://learn.microsoft.com/en-us/graph/api/driveitem-follow) |
| Permanent delete | **Implemented** | **`word permanent-delete`** | Irreversible where policy allows |
| Full Word product (mail merge, macros, compare) | — | — | **Out of scope** — not Graph |

### Word Graph API watchlist

Re-run when Microsoft ships new **drive-item** Word APIs (same cadence as PowerPoint below):

```bash
bash ~/.cursor/skills/msgraph/scripts/run.sh openapi-search --query "driveItem word" --limit 25
bash ~/.cursor/skills/msgraph/scripts/run.sh openapi-search --query "word document" --limit 15
```

Expect **few or no** drive-scoped Word processing paths today; **`workbook`** hits are Excel-only.

## PowerPoint (.pptx) drive-hosted Graph — coverage matrix

**Definition of done (Graph “feature complete” for PowerPoint here):** same as Word: **`powerpoint`** mirrors the per-item **`files`** verbs below; folder/root ops use **`files`**. Graph does not publish a `…/presentation/slides/…` tree on drive items today.

| Graph area (drive item) | `files` | `powerpoint` | `graph invoke` / other |
| --- | --- | --- | --- |
| List / search / delta / shared-with-me | **Implemented** | — | — |
| Metadata (`GET …/items/{id}`) | **Implemented** | **`powerpoint meta`** | — |
| Download content | **Implemented** | **`powerpoint download`** | — |
| Thumbnails (`GET …/items/{id}/thumbnails`) | **`files thumbnails`** | **`powerpoint thumbnails`** | — |
| Upload / large upload / delete | **Implemented** | **`powerpoint upload`**, **`powerpoint upload-large`**, **`powerpoint delete`** | — |
| Copy / move | **Implemented** | **`powerpoint copy`**, **`powerpoint move`** | — |
| Share link / collab / checkout / checkin | **Implemented** | **`powerpoint share`**, **`powerpoint checkout`**, **`powerpoint checkin`** | — |
| Invite / permissions / PATCH permission | **Implemented** | **`powerpoint invite`**, **`powerpoint permissions`**, **`powerpoint permission-remove`**, **`powerpoint permission-update`** | — |
| Versions / restore | **Implemented** | **`powerpoint versions`**, **`powerpoint restore`** | — |
| Convert format | **Implemented** | **`powerpoint convert`** | — |
| Analytics | **Implemented** | **`powerpoint analytics`** | — |
| Per-item activity feed | **Implemented** | **`powerpoint activities`** | — |
| Preview session (`POST …/preview`) | — | **`powerpoint preview`** (also **`word preview`**, **`files preview`**) | — |
| In-deck slides / shapes / slide comments | — | — | **Gap** — no **`…/presentation/…`** path in OpenAPI index; round-trip edit → **`powerpoint download`** → local / OOXML → **`powerpoint upload`** ([`AGENT_WORKFLOWS.md`](./AGENT_WORKFLOWS.md)) |
| Sensitivity / MIP, retention, listItem, follow, permanent-delete | **Implemented** | Same **`word`** subcommand names on **`powerpoint`** | See Word matrix rows above |
| Full PowerPoint product (designer, presenter view) | — | — | **Out of scope** — not Graph |

## PowerPoint Graph API watchlist

When Microsoft adds **drive-item** presentation APIs (e.g. paths under `…/drive/items/{id}/presentation/…`), re-run a local OpenAPI audit and wrap new stable or beta endpoints using the same pattern as [`src/lib/graph-excel-comments-client.ts`](../src/lib/graph-excel-comments-client.ts) + subcommands.

**Suggested check** (requires the [msgraph Cursor skill](https://github.com/merill/graph-skills) or your own Graph OpenAPI export):

```bash
bash ~/.cursor/skills/msgraph/scripts/run.sh openapi-search --query "driveItem presentation" --limit 25
bash ~/.cursor/skills/msgraph/scripts/run.sh openapi-search --query "items workbook" --limit 5
```

From the repo root, **`node scripts/graph-powerpoint-openapi-watch.mjs`** runs equivalent searches (exits **0** with a hint if the skill is missing).

Compare: **`workbook`** paths should return many hits; **`presentation`** on **drive items** should stay empty until Microsoft ships deck APIs. Also scan [Graph API changelog](https://learn.microsoft.com/en-us/graph/whats-new-overview) for “PowerPoint” / “presentation” on **files** / **driveItem**.

## Word / PowerPoint (drive items) — summary

| Graph area | CLI | Notes |
| --- | --- | --- |
| Preview session | **Implemented** | **`word preview`**, **`powerpoint preview`** — POST …/drive/items/{id}/preview (same Graph API; format support is service-dependent); same **`--user`** / **`--drive-id`** / **`--site-id`** / **`--library-drive-id`** as **`files`**. |
| Item metadata + download | **Implemented** | **`word meta`**, **`word download`**, **`powerpoint meta`**, **`powerpoint download`** — GET …/drive/items/{id}; download aligns with **`files download`**. |
| Thumbnails | **Implemented** | **`files thumbnails`**, **`word thumbnails`**, **`powerpoint thumbnails`** — GET …/drive/items/{id}/thumbnails (sizes preauthenticated per Graph). |
| Upload / share / permissions / versions / convert / analytics / activities / checkout | **Implemented** | Same Graph requests as **`files`** — use **`word …`** or **`powerpoint …`** subcommands (see matrices above). |
| listItem / follow / MIP sensitivity / retention / permanentDelete | **Implemented** | **`files …`** and **`word`/`powerpoint` …`** — see Word matrix (**`list-item`**, **`follow`**, **`sensitivity-assign`**, **`retention-label`**, …). |
| In-document review / slide comments (non-Excel) | **Gap** | No **`…/presentation/…/comments`** (or Word equivalent to Excel **`…/workbook/comments`**) in Graph for drive items. Use **`graph invoke`** if Microsoft documents a path, Office client, or OOXML tooling outside this CLI. |


## OneNote

| Graph area | CLI | Notes |
| --- | --- | --- |
| Notebooks, sections, pages, HTML, PATCH | **Implemented** | `onenote` |
| Copy page/section, operations poll | **Implemented** | `onenote copy-page`, `section copy-to-*`, `onenote operation` |
| **GET …/pages/{id}/content** | **Implemented** | `onenote content`, `export`; **`--include-ids`** for `includeIDs=true` |
| **GET …/resources/{id}/content** (binary) | **Implemented** | **`onenote resource-download`** — resource ids from page HTML |
| Page **resources** upload / multipart | **Implemented** | **`onenote create-page-multipart`**, **`onenote patch-page-content-multipart`** |

---

## Files, Teams, To Do, Contacts, Planner, etc

| Graph area | CLI | Notes |
| --- | --- | --- |
| Drive / SharePoint (subset) | **Implemented** | **`files`** — list/search/**`delta`**/**`shared-with-me`**/**`thumbnails`**/**`list-item`**/**`follow`**/**`unfollow`**/**`sensitivity-assign`**/**`sensitivity-extract`**/**`retention-label`**/**`retention-label-remove`**/**`permanent-delete`**/upload/download/**`copy`**/**`move`**/delete/share/**`invite`**/**`permissions`**/**`permission-remove`**/**`permission-update`**/versions/restore/**`checkout`**/**`checkin`**/convert/analytics/**`activities`**/**`preview`**; **`word`/`powerpoint`** mirror per-item **`files`** verbs; targets **`/me/drive`**, **`/users/{id}/drive`**, **`/drives/{id}`**, **`/sites/{id}/drive`** (+ **`--library-drive-id`**); **`sharepoint`** **`resolve-site`**/**`get-site`**/**`drives`**/**`lists`**/**`get-list`**/**`columns`**/**`items`** (default all pages; **`--filter`**/**`--orderby`**/**`--top`**/**`--url`**/**`--all-pages`**), **`get-item`**/**`delete-item`**/**`items-delta`**; **`site-pages`** |
| Excel workbook (worksheets, range, tables, pivots, charts, application, comments) | **Implemented** | **`excel`** — worksheet CRUD; range read/**patch**/**range-clear**; **used-range**; **tables** CRUD + rows add/**patch**/**delete** + **columns** list/get/patch; **pivot-tables** + **pivot-table-*** + **pivot-tables-refresh-all**; **names** + **name-get** + **worksheet-names** / **worksheet-name-get**; **charts** + create/patch/delete; **workbook-get**; **application-calculate**; **session-create** / **session-refresh** / **session-close**; optional **`--session-id`** on mutating calls; **`excel comments-*`** (Graph **beta**). Images/shapes/deep range methods → **`graph invoke`** ([`CLI_REFERENCE.md`](./CLI_REFERENCE.md)) |
| Teams (joined teams, channels, messages, tabs, chats, apps) | **Implemented** | **`teams`** — **`list --user`**; **`channel-files-folder`**; messages/replies/patch/delete/reactions; **`activity-notify`**; **`chat-create`**, **`chat-member-add`**; **`team-member-add`**, **`channel-member-add`**; **tabs** list + **tab-***; **app-catalog** / **app-catalog-get**, **apps** / **app-*** on team, **chat-app-***, **user-app-***; **`teams chats`** = **`/me/chats`** only; meeting lifecycle / RSC / admin → **`graph invoke`** |
| Manager / direct reports / profile / subtree | **Implemented** | **`org manager`**, **`org direct-reports`**, **`org user`**, **`org transitive-reports`** — optional **`--user`** for another user (**`User.Read`** / **`User.Read.All`** per scenario); see **[`PERSONAL_ASSISTANT_DELEGATION.md`](./PERSONAL_ASSISTANT_DELEGATION.md)** |
| **People** (`/me/people`, `/users/{id}/people`) | **Implemented** | **`people list`** (optional **`--search`**, **`--top`**, **`--user`**), **`people get`** |
| **Places** (room lists, rooms, GET place) | **Implemented** | **`rooms`** (**lists**, **rooms**, **find** + **`--query`**, **get**), **`find --rooms`** |
| Bookings | **Implemented** | **`bookings`** — business create/delete/publish/unpublish, currencies list+get, full appointment/customer/service/staff/custom-question CRUD + **calendar-view**, **appointment-cancel**, **staff-availability** (POST; **app-only** token) |
| Bookings **getStaffAvailability** | **Partial (app-only)** | Microsoft documents **no delegated** access — **`bookings staff-availability`** accepts **`--token`** with an application token; delegated **`graph invoke`** will fail |
| Presence | **Implemented** | **`presence`** — **me**, **user**, **bulk**, **set-me** / **set-user**, **clear-me** / **clear-user**, **status-message-set**, **preferred-set** / **preferred-clear**, **clear-location** |
| Raw REST + JSON `$batch` | **Partial** | **`graph invoke`**, **`graph batch`** — escape hatch for any JSON Graph API |
| Online meetings | **Implemented** | `meeting` |
| To Do | **Implemented** | **`todo`** — task CRUD + **`todo delta`** + **`todo lists-delta`** (`lists/delta()`); long tail → **`graph invoke`** |
| Contacts + delta / photo + merge suggestions settings | **Implemented** | `contacts`; **`contacts merge-suggestions`** (Graph beta userSettings) |
| Planner | **Implemented** | **`planner`** incl. beta **`plan-archive`** / **`plan-unarchive`**; **`planner delta`** |
| Search (query) | **Implemented** | **`graph-search`** — presets / **`--types`**, **`--merge-json-file`**, **`--body-file`**, and flags for **`fields`**, **`contentSources`**, **`region`**, **`aggregationFilters`**, **`sortProperties`** (via **`--sort-json-file`**), **`enableTopResults`**, **`trimDuplicates`**; **`find`** stays directory/people |
| Cloud communications (Graph voice/call control; **not** PSTN product scope) | **Gap** | **PSTN / Teams Phone** → **out of scope** (see [`GRAPH_INVOKE_BOUNDARIES.md`](./GRAPH_INVOKE_BOUNDARIES.md)). Other APIs → **`graph invoke`** if org-approved. |

---

## Discovery & Insights

| Graph area | CLI | Notes |
| --- | --- | --- |
| **Office Insights — `/me/insights/trending`** | **Implemented** | **`insights trending`** — documents trending around the user; **`--user`** for delegation. |
| **Office Insights — `/me/insights/used`** | **Implemented** | **`insights used`** — documents the user has used recently. |
| **Office Insights — `/me/insights/shared`** | **Implemented** | **`insights shared`** — documents shared with the user. |
| **`GET /me/drive/recent`** | **Implemented** | **`files recent`** — Office.com "Recent" rail. |
| **`GET /drives/{id}/items/{id}/activities`** | **Implemented** | **`files activities <fileId>`** — per-item activity feed; honors **`--user`** / **`--drive-id`** / **`--site-id`** / **`--library-drive-id`**. |
| **`POST /me/drive/items/{id}/preview`** | **Implemented** | **`files preview <fileId>`** — preview session for any drive item; complements **`word preview`** / **`powerpoint preview`** for non-Office types. |
| **`GET /me/followedSites`** + **add/remove** | **Implemented** | **`sharepoint followed-sites`** / **`sharepoint follow <siteId>`** / **`sharepoint unfollow <siteId>`** — SharePoint "Following" rail. |

---

## Outlook Groups (Microsoft 365 groups)

| Graph area | CLI | Notes |
| --- | --- | --- |
| **`GET /me/memberOf`** filtered to Microsoft 365 groups | **Implemented** | **`groups list`** — lists Outlook / Microsoft 365 groups the user belongs to (`groupTypes` includes `Unified`). |
| **`GET /groups/{id}/conversations`** | **Implemented** | **`groups conversations <groupId>`** — list group conversations. |
| **`GET /groups/{id}/conversations/{id}/threads`** | **Implemented** | **`groups thread <groupId> <conversationId>`** — list threads. |
| **`GET /groups/{id}/conversations/{id}/threads/{id}/posts`** | **Implemented** | **`groups posts <groupId> <conversationId> <threadId>`** — list posts. |
| **`POST .../posts/{id}/reply`** | **Implemented** | **`groups post-reply`** — reply to a group thread post. Reuses **`Group.ReadWrite.All`**. |

---

## Approvals (Teams Approvals app)

| Graph area | CLI | Notes |
| --- | --- | --- |
| **`GET /me/approvals`** | **Implemented** | **`approvals list`** — Approvals visible to the user (Teams Approvals app, Power Automate approvals). Beta. Paging: **`--all`** (follow `@odata.nextLink`), **`--next <url>`**, JSON includes **`@odata.nextLink`** when present. |
| **`GET /me/approvals/{id}`** | **Implemented** | **`approvals get`** |
| **`GET /me/approvals/{id}/steps`** | **Implemented** | **`approvals steps`** |
| **`PATCH /me/approvals/{id}/steps/{stepId}`** | **Implemented** | **`approvals respond <id> <stepId> --decision approve\|deny --justification "<text>"`** |
| **`DELETE /me/approvals/{id}`** (owner cancel) | **Implemented** | **`approvals cancel <id>`** — optional **`--if-match`**; otherwise the CLI **`GET`**s the approval once for **`@odata.etag`**. |

Scope: delegated **`ApprovalSolution.ReadWrite`** (canonical name; identifier `6768d3af-4562-48ff-82d2-c5e19eb21b9c`). A narrower **`ApprovalSolutionResponse.ReadWrite`** exists for read-and-respond only. Added to [`graph-oauth-scopes.ts`](../src/lib/graph-oauth-scopes.ts).

---

## Online meeting recordings & transcripts (delegated)

| Graph area | CLI | Notes |
| --- | --- | --- |
| **`GET /me/onlineMeetings/{id}/recordings`** | **Implemented** | **`meeting recordings <meetingId>`** — list call recordings on a single meeting. |
| **`GET …/recordings/{id}/content`** | **Implemented** | **`meeting recording-download <meetingId> <recordingId> [--out <path>]`** — recording bytes. |
| **`getAllRecordings(meetingOrganizerUserId,start,end)`** + delta | **Implemented** | **`meeting recordings-all --start <iso> --end <iso> [--user] [--delta] [--state-file <path>]`**. |
| **`GET /me/onlineMeetings/{id}/transcripts`** | **Implemented** | **`meeting transcripts <meetingId>`** |
| **`GET …/transcripts/{id}/content`** | **Implemented** | **`meeting transcript-download`** — VTT body; pass **`--format text`** to also fetch `metadataContent`. |
| **`getAllTranscripts(...)`** + delta | **Implemented** | **`meeting transcripts-all`** mirrors recordings shape. |

Scopes: **`OnlineMeetingRecording.Read.All`** (new) + **`OnlineMeetingTranscript.Read.All`** (existing). 403 typically means tenant Stream/Teams policy; not a CLI bug.

---

## Teams activity feed (delegated)

| Graph area | CLI | Notes |
| --- | --- | --- |
| **`POST /me/teamwork/sendActivityNotification`** | **Implemented** | **`teams activity-notify`** — user-targeted ping (the bell). |
| **`POST /chats/{id}/sendActivityNotification`** | **Implemented** | **`teams activity-notify --chat-id <id>`** — chat-scoped notification. |
| **`POST /users/{id}/teamwork/sendActivityNotification`** | **Implemented** | **`teams activity-notify --user-id <id>`** (typically **`--token`** with application access token); see [`GRAPH_INVOKE_BOUNDARIES.md`](./GRAPH_INVOKE_BOUNDARIES.md). |

Scope: delegated **`TeamsActivity.Send`**.

---

## Subscriptions & change notifications

| Graph area | CLI | Notes |
| --- | --- | --- |
| Create/delete subscription | **Implemented** | `subscribe` |
| Webhook validation | **Implemented** | `webhook-server` helper |
| List / renew subscriptions | **Implemented** | **`subscribe list`**, **`subscribe renew`**, **`subscriptions renew-all`** (plus create/cancel) — run **`renew-all`** on a schedule (e.g. cron) before expiry |

---

## What “new” Graph features usually mean here

1. **Delta + long-running sync** — **`contacts delta`**, **`todo delta`**, **`planner delta`**, **`outlook-graph messages-delta`**, **`graph-calendar events-delta`** support **`--state-file`** where listed; else **`--next`** / **`--url`**.
2. **Microsoft Search** — **`graph-search`** with **`--preset default`**, **`extended`**, **`connectors`**, or explicit **`--types`**; exotic verticals may need **`graph invoke`**.
3. **OneNote** — advanced ink scenarios beyond multipart HTML + binary parts may need Graph directly.
4. **PowerPoint (and Word) deck/body APIs** — if OpenAPI gains **`…/items/{id}/presentation/…`** (or similar) with stable contracts, add **`graph-*-client`** modules and **`powerpoint`** / **`word`** subcommands; until then **`word`**/**`powerpoint`** wrap preview/meta/download/thumbnails **and** mirror **`files`** per-item lifecycle — see **PowerPoint Graph API watchlist** above.

---

*Last updated: 2026-05-05 — **Teams** app catalog + installs; **Presence** extended APIs. **Drive** **`files`** list-item / follow / MIP sensitivity / retention / permanent-delete (+ **`word`/`powerpoint`** mirrors). **Directory / people / rooms**, **Word**/**PowerPoint**, **Excel** closure as documented. Scopes in [`GRAPH_SCOPES.md`](./GRAPH_SCOPES.md).*
