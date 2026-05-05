# Graph product parity matrix (m365-agent-cli)

**Purpose:** Map major **Microsoft 365 workloads** to **CLI commands**, Graph coverage (**Implemented** / **Partial** / **Gap**), and notes (delegation, app-only, escape hatch).

**Related:** [`GRAPH_API_GAPS.md`](./GRAPH_API_GAPS.md) (detailed API-by-area gaps + **closure targets** table), [`GRAPH_WRAPPER_GAP_AUDIT.md`](./GRAPH_WRAPPER_GAP_AUDIT.md) (prioritized gap backlog, beta inventory, OpenAPI strict gate, pagination/delegation notes), [`WORD_POWERPOINT_EDITING.md`](./WORD_POWERPOINT_EDITING.md) (Word/PowerPoint edit workflows, OOXML, checkout/convert), [`GRAPH_SCOPES.md`](./GRAPH_SCOPES.md), [`GRAPH_PERMISSION_FEATURE_MATRIX.md`](./GRAPH_PERMISSION_FEATURE_MATRIX.md) (which **Entra permissions** map to which CLI feature areas), [`GRAPH_INVOKE_BOUNDARIES.md`](./GRAPH_INVOKE_BOUNDARIES.md) (invoke-only surfaces + recipes), [`CLI_REFERENCE.md`](./CLI_REFERENCE.md) (command inventory), [`AGENT_WORKFLOWS.md`](./AGENT_WORKFLOWS.md) (orchestration patterns), [`CLI_SCRIPTING_APPENDIX.md`](./CLI_SCRIPTING_APPENDIX.md) (`--json` / read-only inventory).

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
| **To Do** | `todo` | **Implemented** | **`todo delta`** + **`todo lists-delta`** support **`--state-file`** (tasks vs task **lists**). **`todo attachment-session`** + **`todo root`** cover **`…/attachmentSessions`** and **`…/todo`** (delete gated with **`--confirm`**). |
| **Planner** | `planner` | **Implemented** | **`planner delta`** + **`--state-file`**; beta **`plan-archive`** / **`plan-unarchive`**. |
| **Contacts** | `contacts` | **Implemented** | **`contacts delta`** supports **`--state-file`**. **`contacts extension`** supports **`-f`/`--folder`** + **`--child-folder`** for nested folder extension URLs. |
| **OneNote** | `onenote` | **Implemented** | |
| **Excel (workbook on drive)** | `excel` | **Implemented** | Worksheets, range read/patch/**range-clear**, used-range, **tables** CRUD + rows add/patch/delete + **columns** list/get/patch, **pivot-tables** lifecycle + refresh, **names** list + **name-get** + **worksheet-names** / **worksheet-name-get**, charts, **workbook-get**, **application-calculate**, **session-create** / **session-refresh** / **session-close** (optional **`--session-id`** on mutating calls), **`excel comments-*`** (Graph **beta**); images/shapes/deep range actions → **`graph invoke`**. |
| **Word / PowerPoint (files)** | `word`, `powerpoint`, `files` | **Implemented** | **`word`/`powerpoint`** mirror all **per-item** **`files`** Graph APIs including **`list-item`**, **`follow`/`unfollow`**, **`sensitivity-assign`/`extract`**, **`retention-label`** (+ **`remove`**), **`permanent-delete`**, checkout/share/versions/upload/… **Platform gaps only:** no Graph **in-file** Word/slide **comment** threads or **presentation object model** (slides/shapes) like Excel **`workbook/…`**; use **`graph invoke`** when Microsoft documents beta paths, or local OOXML/desktop Office. **Folder** **list**/**delta**/**search** → **`files`**. |
| **OneDrive / SharePoint files** | `files`, `sharepoint`, `site-pages` | **Implemented** | **`files`** — list/search/**`delta`**/**`shared-with-me`**/**`thumbnails`**/upload/download/**`copy`**/**`move`**/delete/share/**`invite`**/**`permissions`**/**`permission-remove`**/**`permission-update`**/versions/**`checkout`**/**`checkin`**/convert/analytics + drive flags; **`sharepoint`** **`resolve-site`**, **`get-site`**, **`drives`**, **`lists`**, **`get-list`**, **`columns`**, **`items`** (OData **`--filter`** / **`--orderby`** / **`--top`** / **`--url`** / **`--all-pages`**), **`get-item`**, **`create-item`** / **`update-item`** (**`--json-file`**), **`delete-item`**, **`items-delta`**; followed sites; **`site-pages`**. Tenant admin / unsupported list APIs → **`graph invoke`**. |
| **Teams** | `teams` | **Implemented** | Joined teams (**`list --user`**), channels, **channel-files-folder**, messages/replies/patch/delete/reactions, **tabs**, **members**, **app-catalog** / catalog get, **apps** on team (list/get/add/patch/upgrade/delete), **chat-apps** lifecycle, **user-apps** (personal scope, optional **`--user`**), chats (**`/me/chats`** only), **`activity-notify`** (incl. **`--user-id`** for **`/users/{id}/teamwork/sendActivityNotification`**, usually app token), etc. Meeting lifecycle / tenant admin / RSC-only surfaces → **`graph invoke`**. |
| **Presence** | `presence` | **Implemented** | Read (**`me`**, **`user`**, **`bulk`**), session **set/clear**, **status-message-set**, **preferred-set** / **preferred-clear**, **clear-location**; subscriptions via **`subscribe`** + **`serve`**. |
| **Bookings** | `bookings` | **Implemented** | Full Graph v1 Bookings under `/solutions/bookingBusinesses` (incl. **business-create** / **business-delete** / **business-publish** / **business-unpublish**, **currency-get**, **custom-question**); **`staff-availability`** remains **app-only** per Microsoft (`--token`). |
| **Microsoft Search** | `graph-search`, `find` | **Implemented** (`graph-search`) | Presets / **`--types`**, **`searchRequest`** shaping (**`--merge-json-file`**, **`--fields`**, **`--content-sources`**, **`--region`**, **`--aggregation-filters`**, **`--sort-json-file`**, **`--enable-top-results`**, **`--trim-duplicates`**), or raw **`--body-file`** for multi-request POST bodies. Directory **`find`** unchanged. |
| **Copilot (Graph)** | `copilot` | **Implemented** | Full `/copilot` OpenAPI surface: root + communications + tenant `interactionHistory` nav, retrieval/search, conversations/messages (+ `$count`), OData chat actions, agents (+ `$count`), settings (+ deletes), `reports` usage functions + nav get/patch/delete, admin nav/catalog/settings (+ deletes), packages (+ `$count`, zip get/put/delete), meeting insights CRUD + `$count`, per-user/tenant interaction export, **`ai-user`** (`/copilot/users` CRUD, interactionHistory, onlineMeetings), **`activity-feed`** (+ root patch/delete, `$count`); preview/beta labeled in help. |
| **Subscriptions / webhooks** | `subscribe`, `subscriptions`, `serve` | **Implemented** | **`subscriptions renew-all`** for automation. |
| **Directory / people / rooms** | `find`, `people`, `org`, `rooms`, `schedule`, `suggest` | **Implemented** | **`org user`** / **`org transitive-reports`** + manager/direct-reports; **`people list|get`** (`/me/people`, optional **`--user`**); **`find --rooms`**; **`rooms get`** / **`rooms find --query`**; admin directory still via **`graph invoke`**. |
| **Discovery / Insights** | `insights`, `files recent`, `files activities`, `files preview`, `sharepoint followed-sites`/`follow`/`unfollow` | **Implemented** | Office Insights (`/me/insights/trending`/`used`/`shared`), drive recent + per-item activity feed, generic drive item preview, SharePoint Following rail. |
| **Viva / employee experience (Graph beta)** | `viva` | **Implemented** | User scope + tenant **`/employeeExperience`** (singleton, **communities**, **async ops**, **goals** + export jobs + content, **learning** root + **providers** + **contents** + provider activities, **roles** + members + nested **user** / **mailboxSettings** / **serviceProvisioningErrors**, community **owners** + **UPN** lookup + owner **mailboxSettings** / **serviceProvisioningErrors**), **admin/org itemInsights**, **`workHoursAndLocations`** (occurrences / recurrences / view / setCurrentLocation), **meeting Engage** (conversations, messages, replies, reactions, navigation). Unusual **`$expand`** / undocumented preview → **`graph invoke --beta`**. |
| **Outlook Groups** | `groups list`, `conversations`, `thread`, `posts`, `post-reply` | **Implemented** | Microsoft 365 group conversation surface (Outlook Groups inbox); reuses **`Group.ReadWrite.All`**. |
| **Approvals** | `approvals list`, `get`, `steps`, `respond` | **Implemented** | Teams Approvals app + Power Automate approvals via **`/me/approvals`** (beta). New scope **`ApprovalSolution.ReadWrite`** (canonical delegated). |
| **Meeting recordings / transcripts (delegated)** | `meeting recordings`/`recording-download`/`recordings-all`, `meeting transcripts`/`transcript-download`/`transcripts-all` | **Implemented** | Per-meeting + tenant-wide `getAllRecordings` / `getAllTranscripts` (+ delta with `--state-file`). 403 = tenant Stream/Teams policy, not CLI bug. |
| **Teams activity notifications** | `teams activity-notify` | **Implemented** | **`POST /me/teamwork/sendActivityNotification`**, **`POST /chats/{id}/sendActivityNotification`**, and **`POST /users/{id}/teamwork/sendActivityNotification`** via **`--user-id`** (typically **`--token`** application access token). |
| **Teams Phone / PSTN** | — | **Out of scope** | Not documented or wrapped; use Teams admin center / carrier tooling — [`GRAPH_INVOKE_BOUNDARIES.md`](./GRAPH_INVOKE_BOUNDARIES.md). |
| **Other Cloud Communications** (Graph call control, call records, …) | — | **Gap** | **`graph invoke`** if org-approved; PSTN product scenarios are **out of scope** — same doc. |

---

## Escape hatch

Any API not wrapped: **`m365-agent-cli graph invoke`** and **`graph batch`** with appropriate delegated scopes in Entra.

---

*Last updated: 2026-05-05 — **Viva** **`viva`**: **Implemented** (tenant `/employeeExperience` including owner/member deep navigation, admin/org insights, work hours & locations, meeting Engage messages/replies/reactions, plus user work time / insights / roles / learning). **Copilot** **`copilot`**: **Implemented** (conversations/messages/agents/settings/admin/packages/zip, usage reports, meeting insights, per-user + tenant interaction export, **activity-feed**; chat uses `microsoft.graph.copilot.*` actions). **Teams** + **Presence**: **Implemented** (app catalog, team/chat/personal app installs, status message / preferred presence / **clear-location**). **Directory / people / rooms**, Bookings, **`graph-search`**, OneDrive / SharePoint, Word/PowerPoint, Excel **`excel`**, **`scripts/graph-powerpoint-openapi-watch.mjs`**, Phase 1–3 closures — see [`GRAPH_API_GAPS.md`](./GRAPH_API_GAPS.md).*

