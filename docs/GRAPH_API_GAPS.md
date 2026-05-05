# Microsoft Graph vs this CLI — capabilities and gaps

**Purpose:** Track **Graph API areas** that are **implemented** in `m365-agent-cli`, **partially** covered, or **not** exposed, so we can prioritize work and set expectations.

**Related:** [`MIGRATION_TRACKING.md`](./MIGRATION_TRACKING.md), [`GRAPH_EWS_PARITY_MATRIX.md`](./GRAPH_EWS_PARITY_MATRIX.md), [`GRAPH_SCOPES.md`](./GRAPH_SCOPES.md), [`GRAPH_PRODUCT_PARITY_MATRIX.md`](./GRAPH_PRODUCT_PARITY_MATRIX.md) (workloads ↔ commands), [`GRAPH_INVOKE_BOUNDARIES.md`](./GRAPH_INVOKE_BOUNDARIES.md) (invoke-only surfaces).

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
| **Word (.docx)** | **`word`**: **`preview`**, **`meta`**, **`download`**, **`thumbnails`** (+ drive location flags like **`files`**) | — | **Graph-complete** for drive-hosted Word: every stable delegated drive-item op for `.docx` is **`files`** or **`word`** per the Word matrix below, or **Gap** (in-document comments — no Graph path in OpenAPI index). |
| **PowerPoint (.pptx)** | **`powerpoint`**: **`preview`**, **`meta`**, **`download`**, **`thumbnails`** (+ drive flags) | — | **Graph-complete** for drive-hosted decks: Microsoft Graph exposes no `…/presentation/…` object model on drive items (unlike Excel `…/workbook/…`); **`powerpoint`** wraps preview/meta/download/thumbnails only. Slide-level / in-deck comments → **Gap** until Graph ships APIs; use **`files`** for permissions/versions/**`delta`**/**`copy`**/**`invite`**/etc. and **`graph invoke`** for beta experiments — see **[`GRAPH_INVOKE_BOUNDARIES.md`](./GRAPH_INVOKE_BOUNDARIES.md)** and **PowerPoint Graph API watchlist** below. |
| **Teams** | **`teams`** — joined teams, channels, messages (send/read/**patch**/**soft-delete**/**hard-delete**/**undo-soft** with **`--beta`** where needed), replies, reactions, **`activity-notify`**, **`chat-create`** / **`chat-member-add`**, **`team-member-add`** / **`channel-member-add`**, **tabs** (list/**get**/**create**/**patch**/**delete**), chats, apps | RSC / tenant admin | **`teams list --user`** = **`GET /users/{id}/joinedTeams`**; chat **list** stays **`/me/chats`**; **`POST /users/{id}/teamwork/sendActivityNotification`** app-only → **`graph invoke`**; admin/RSC/shifts → **[`GRAPH_INVOKE_BOUNDARIES.md`](./GRAPH_INVOKE_BOUNDARIES.md)**. |
| **Presence** | **`presence`** read/set/clear | — | All **direct** presence Graph reads in scope; **subscription** automation via **`subscribe`** + **`serve`** (see command help). |
| **Search** | **`graph-search`**: `--preset default|extended|connectors` or `--types` | Exotic connectors | Presets cover main **entityTypes**; long tail **invoke** recipes in gaps + boundaries doc. |
| **Files / SharePoint** | **`files`** (delta, **`shared-with-me`**, **`thumbnails`**, copy/move, drive flags, **`invite`**, **`permissions`**, **`permission-remove`**, **`permission-update`**), **`sharepoint`** (get/delete item, items-delta, **`resolve-site`**, **`--json-file`** on create/update), **`site-pages`**, **`excel`** (workbook object model incl. comments beta), **`teams channel-files-folder`** | — | Core drive + list sync + sharing + Excel on-item workbook APIs wrapped; long tail stays **`graph invoke`**. |
| **Copilot** | **`copilot`** | — | Stable Graph `/copilot` endpoints only; preview APIs labeled in help. |
| **Directory / rooms** | **`find`**, **`rooms`**, **`schedule`**, **`suggest`** | Admin directory | Delegated **Places/People** paths per **GRAPH_SCOPES**; admin-only → **invoke**. |
| **Bookings** | Full CRUD + reads | **`bookings staff-availability`** | Delegated token **fails by design** — use **`--token`** app-only; document, no “delegated fix”. |
| **Cloud Communications** | — | **All** | **[`GRAPH_INVOKE_BOUNDARIES.md`](./GRAPH_INVOKE_BOUNDARIES.md)** **invoke** recipes only. |
| **Mail migration** | **`oof`**, **`rules`**, **`update-event`** (Graph id) | **`auto-reply`** (EWS) | **auto-reply** 1:1 Graph template UX = **not** a goal; see [`MIGRATION_TRACKING.md`](./MIGRATION_TRACKING.md). |
| **Discovery / Insights** (Phase 1) | **`insights trending`** / **`used`** / **`shared`**, **`files recent`** / **`files activities`** / **`files preview`**, **`sharepoint followed-sites`** / **`follow`** / **`unfollow`** | — | Wraps **`/me/insights/*`**, **`/me/drive/recent`**, **`/drives/{id}/items/{id}/activities`**, **`/me/drive/items/{id}/preview`**, **`/me/followedSites`** (+`add`/`remove`). Reuses **`Sites.ReadWrite.All`** + **`Files.ReadWrite.All`**. |
| **Outlook Groups** (Phase 2) | **`groups list`** / **`conversations`** / **`thread`** / **`posts`** / **`post-reply`** | — | Wraps **`/me/memberOf`** filtered to Microsoft 365 groups + **`/groups/{id}/conversations/threads/posts/reply`**. Reuses **`Group.ReadWrite.All`**. |
| **Approvals** (Phase 2) | **`approvals list`** / **`get`** / **`steps`** / **`respond`** | — | Wraps **`/me/approvals`** + **`/steps`** (beta). Adds delegated **`ApprovalSolution.ReadWrite`** scope (canonical name; identifier `6768d3af-4562-48ff-82d2-c5e19eb21b9c`). A narrower **`ApprovalSolutionResponse.ReadWrite`** is available for read-and-respond only. |
| **Meeting recordings & transcripts** (Phase 3) | **`meeting recordings`** / **`recording-download`** / **`recordings-all`** (+ `--delta`); **`meeting transcripts`** / **`transcript-download`** / **`transcripts-all`** | — | Wraps **`/me/onlineMeetings/{id}/recordings`** + **`/transcripts`**, **`getAllRecordings(...)`** / **`getAllTranscripts(...)`**, and **`recordings/delta()`** / **`transcripts/delta()`**. Adds **`OnlineMeetingRecording.Read.All`** (transcripts already have **`OnlineMeetingTranscript.Read.All`**). 403 typically means tenant Stream/Teams policy. |
| **Teams activity feed** (Phase 3) | **`teams activity-notify`** | App-only `users/{id}/teamwork/sendActivityNotification` stays **`graph invoke`** | Wraps delegated **`POST /me/teamwork/sendActivityNotification`** + **`POST /chats/{id}/sendActivityNotification`**. Adds **`TeamsActivity.Send`**. |

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

**Definition of done (Graph “feature complete” for Word here):** every **stable delegated** drive-item operation that applies to `.docx` is **Implemented** via **`files`** (or **`word`** where we expose an Office-oriented entry), or explicitly **Gap / invoke** below. **`word`** is intentionally thin: preview + convenience mirrors; avoid duplicating every **`files`** subcommand under **`word`** (see [CLI_REFERENCE.md](./CLI_REFERENCE.md) Word section).

| Graph area (drive item) | `files` | `word` | `graph invoke` / other |
| --- | --- | --- | --- |
| List / search / delta / shared-with-me | **Implemented** | — | — |
| Metadata (`GET …/items/{id}`) | **Implemented** | **`word meta`** | — |
| Download content | **Implemented** | **`word download`** | — |
| Thumbnails (`GET …/items/{id}/thumbnails`) | **`files thumbnails`** | **`word thumbnails`** (same impl as **`powerpoint thumbnails`**) | — |
| Upload / large upload / delete | **Implemented** | — | — |
| Copy / move | **Implemented** | — | — |
| Share link / collab / checkout / checkin | **Implemented** | — | — |
| Invite / permissions / PATCH permission | **Implemented** | — | — |
| Versions / restore | **Implemented** | — | — |
| Convert format | **Implemented** | — | — |
| Analytics | **Implemented** | — | — |
| Preview session (`POST …/preview`) | — | **`word preview`** (also **`powerpoint preview`**) | **`graph invoke`** if you need non-CLI automation only |
| In-document comments / review (Word-specific) | — | — | **Gap** — no **`…/workbook/comments`–style** Word path in OpenAPI index; Office client / OOXML / beta docs only |
| Sensitivity / MIP labels (item) | Partial | — | Often **`graph invoke`** beta; confirm per tenant docs |
| Full Word product (mail merge, macros, compare) | — | — | **Out of scope** — not Graph |

### Word Graph API watchlist

Re-run when Microsoft ships new **drive-item** Word APIs (same cadence as PowerPoint below):

```bash
bash ~/.cursor/skills/msgraph/scripts/run.sh openapi-search --query "driveItem word" --limit 25
bash ~/.cursor/skills/msgraph/scripts/run.sh openapi-search --query "word document" --limit 15
```

Expect **few or no** drive-scoped Word processing paths today; **`workbook`** hits are Excel-only.

## PowerPoint (.pptx) drive-hosted Graph — coverage matrix

**Definition of done (Graph “feature complete” for PowerPoint here):** every **documented** Microsoft Graph operation that applies to a **drive-hosted `.pptx` as an Office file** (not generic file CRUD) is **Implemented** via **`powerpoint`** or **`files`**, or is explicitly **Gap** because Graph does not publish a comparable surface (no `…/presentation/slides/…` tree in the OpenAPI catalog today). **`powerpoint`** is intentionally thin — same pattern as **`word`**: avoid duplicating every **`files`** subcommand.

| Graph area (drive item) | `files` | `powerpoint` | `graph invoke` / other |
| --- | --- | --- | --- |
| List / search / delta / shared-with-me | **Implemented** | — | — |
| Metadata (`GET …/items/{id}`) | **Implemented** | **`powerpoint meta`** | — |
| Download content | **Implemented** | **`powerpoint download`** | — |
| Thumbnails (`GET …/items/{id}/thumbnails`) | **`files thumbnails`** | **`powerpoint thumbnails`** | — |
| Upload / large upload / delete | **Implemented** | — | — |
| Copy / move | **Implemented** | — | — |
| Share link / collab / checkout / checkin | **Implemented** | — | — |
| Invite / permissions / PATCH permission | **Implemented** | — | — |
| Versions / restore | **Implemented** | — | — |
| Convert format | **Implemented** | — | — |
| Analytics | **Implemented** | — | — |
| Preview session (`POST …/preview`) | — | **`powerpoint preview`** (also **`word preview`**, **`files preview`**) | **`graph invoke`** if you need non-CLI automation only |
| In-deck slides / shapes / slide comments | — | — | **Gap** — no **`…/presentation/…`** path in OpenAPI index; round-trip edit → **`powerpoint download`** → local / OOXML → **`files upload`** ([`AGENT_WORKFLOWS.md`](./AGENT_WORKFLOWS.md)) |
| Sensitivity / MIP labels (item) | Partial | — | Often **`graph invoke`** beta; confirm per tenant docs |
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
| In-document review / slide comments (non-Excel) | **Gap** | Local OpenAPI index does not expose a **`…/presentation/…/comments`** (or Word) drive-item surface comparable to Excel **`…/workbook/comments`**. Use **`graph invoke`** against beta docs if your scenario is supported, or Office client / OOXML automation outside this CLI. |


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
| Drive / SharePoint (subset) | **Partial** | **`files`** — list/search/**`delta`**/**`shared-with-me`**/**`thumbnails`**/upload/download/**`copy`**/**`move`**/delete/share/**`invite`**/**`permissions`**/**`permission-remove`**/**`permission-update`**/versions/restore/checkin/convert/analytics; targets **`/me/drive`**, **`/users/{id}/drive`**, **`/drives/{id}`**, **`/sites/{id}/drive`** (+ **`--library-drive-id`**); **`sharepoint`** lists/items/**`get-item`**/**`delete-item`**/**`items-delta`**/**`resolve-site`**; **`site-pages`** |
| Excel workbook (worksheets, range, tables, pivots, charts, application, comments) | **Implemented** | **`excel`** — worksheet CRUD; range read/**patch**/**range-clear**; **used-range**; **tables** CRUD + rows add/**patch**/**delete** + **columns** list/get/patch; **pivot-tables** + **pivot-table-*** + **pivot-tables-refresh-all**; **names** + **name-get** + **worksheet-names** / **worksheet-name-get**; **charts** + create/patch/delete; **workbook-get**; **application-calculate**; **session-create** / **session-refresh** / **session-close**; optional **`--session-id`** on mutating calls; **`excel comments-*`** (Graph **beta**). Images/shapes/deep range methods → **`graph invoke`** ([`CLI_REFERENCE.md`](./CLI_REFERENCE.md)) |
| Teams (joined teams, channels, messages, tabs, chats) | **Partial** | **`teams`** — **`list --user`**; **`channel-files-folder`**; message send/patch/soft-delete/hard-delete/undo-soft; **`activity-notify`**; **`chat-create`**, **`chat-member-add`**; **`team-member-add`**, **`channel-member-add`**; **tabs** list + **tab-*** CRUD; reactions; **`teams chats`** list **`/me`** only; meeting lifecycle / RSC / admin → **`graph invoke`** |
| Manager / direct reports | **Implemented** | **`org manager`**, **`org direct-reports`** — optional **`--user`** for another user’s hierarchy (**`User.Read`** / **`User.Read.All`** per scenario); see **[`PERSONAL_ASSISTANT_DELEGATION.md`](./PERSONAL_ASSISTANT_DELEGATION.md)** |
| Bookings | **Partial** | **`bookings`** — full CRUD + **staff-availability** (POST; **app-only** token) + **appointment-cancel** |
| Bookings **getStaffAvailability** | **Partial (app-only)** | Microsoft documents **no delegated** access — **`bookings staff-availability`** accepts **`--token`** with an application token; delegated **`graph invoke`** will fail |
| Presence | **Partial** | **`presence`** — **me**, **user**, **bulk**, **set-me** / **set-user** (prints `sessionId`), **clear-me** / **clear-user** |
| Raw REST + JSON `$batch` | **Partial** | **`graph invoke`**, **`graph batch`** — escape hatch for any JSON Graph API |
| Online meetings | **Implemented** | `meeting` |
| To Do | **Implemented** | **`todo`** — task CRUD + **`todo delta`** + **`todo lists-delta`** (`lists/delta()`); long tail → **`graph invoke`** |
| Contacts + delta / photo | **Implemented** | `contacts` |
| Planner | **Implemented** | **`planner`** incl. beta **`plan-archive`** / **`plan-unarchive`**; **`planner delta`** |
| Search (query) | **Partial** | **`graph-search`** — **`--preset extended`** or **`connectors`** (connector-heavy **entityTypes**) or **`--types`**; **`find`** — exotic connectors may still need **`graph invoke`** |
| Cloud communications (calls, PSTN, etc.) | **Gap** | Use **`graph invoke`** or dedicated apps; not wrapped |

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
| **`GET /me/approvals`** | **Implemented** | **`approvals list`** — Approvals visible to the user (Teams Approvals app, Power Automate approvals). Beta. |
| **`GET /me/approvals/{id}`** | **Implemented** | **`approvals get`** |
| **`GET /me/approvals/{id}/steps`** | **Implemented** | **`approvals steps`** |
| **`PATCH /me/approvals/{id}/steps/{stepId}`** | **Implemented** | **`approvals respond <id> <stepId> --decision approve\|deny --justification "<text>"`** |
| Cancel | **Gap** | No first-class action exposed in v1.0/beta as of writing — use **`graph invoke`** if your tenant exposes a `cancel` action. |

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
| **`POST /users/{id}/teamwork/sendActivityNotification`** | **Gap** | App-only path — stays **`graph invoke`**, see [`GRAPH_INVOKE_BOUNDARIES.md`](./GRAPH_INVOKE_BOUNDARIES.md). |

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
4. **PowerPoint (and Word) deck/body APIs** — if OpenAPI gains **`…/items/{id}/presentation/…`** (or similar) with stable contracts, add **`graph-*-client`** modules and **`powerpoint`** / **`word`** subcommands; until then **`powerpoint`** stays preview/meta/download/thumbnails + **`files`** — see **PowerPoint Graph API watchlist** above.

---

*Last updated: 2026-05-05 — **Excel** closure: **`excel`** pivots, table CRUD, columns, row patch/delete, **range-clear**, **workbook-get**, **application-calculate**, **session-refresh**, names / worksheet names, **`--session-id`** on mutating calls. PowerPoint / Word **Graph-complete** closure; PowerPoint matrix + watchlist script; Phase 1–3 closures; scopes in [`GRAPH_SCOPES.md`](./GRAPH_SCOPES.md).*
