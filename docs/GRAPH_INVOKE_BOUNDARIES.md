# Graph API: intentional CLI boundaries (`graph invoke`)

Some Microsoft Graph areas are **not** wrapped as first-class subcommands in `m365-agent-cli`. This document explains why and how to call them safely.

## Teams Phone / PSTN (out of scope)

**PSTN**, **Teams Phone**, direct routing, and telephony administration are **intentionally out of scope** for `m365-agent-cli`. This project does not document recipes, track product gaps, or plan first-class wrappers for those surfaces. Use the **Microsoft Teams admin center**, your **telephony provider** tooling, or **PowerShell** / partner automation instead.

## Cloud Communications (Graph — not wrapped here)

Other **Microsoft Graph Cloud Communications** APIs (for example **call records**, call control, or org-specific voice scenarios) carry **high compliance and consent** requirements. They are **not** covered by this CLI’s first-class commands; if your organization approves automation, use **`graph invoke`** against the current [Graph API reference](https://learn.microsoft.com/en-us/graph/api/overview) and the correct **application** or **delegated** permissions.

> **Carve-out:** **delegated meeting recordings & transcripts** (per-meeting + tenant-wide `getAllRecordings(...)` / `getAllTranscripts(...)` and their `delta` functions) **are** wrapped — see **`meeting recordings*`** / **`meeting transcripts*`** with **`OnlineMeetingRecording.Read.All`** / **`OnlineMeetingTranscript.Read.All`**. Tenant Stream/Teams policies can still return 403; that's policy, not CLI breakage.

## Teams resource-specific consent (RSC) / app-only admin scenarios

Many Teams admin and tenant-wide operations require **application** permissions or policies that **delegated** interactive users cannot satisfy. Prefer **Microsoft Teams admin center**, **PowerShell**, or dedicated automation with app-only tokens—not the default delegated **`login`** flow.

For delegated chat/channel scenarios already covered, use **`teams`**. To script against files shown in a channel’s **Files** tab, use **`teams channel-files-folder`** (with team and channel ids) and then **`files list --drive-id … --folder …`** (wrapped); admin-only paths remain **`graph invoke`**.

### Example (application permission — illustrative only)

```bash
m365-agent-cli graph invoke --token "$APP_TOKEN" -X GET "/groups?\$filter=resourceProvisioningOptions/Any(x:x eq 'Team')&\$top=5"
```

Confirm **`Group.Read.All`** (or **Group.ReadWrite.All**) on the app registration and admin consent.

## Bookings **getStaffAvailability**

Microsoft documents **no delegated** access for some Bookings APIs. The CLI exposes **`bookings staff-availability`** with **`--token`** for an **application** token when required.

Delegated tokens should fail with **403/401** from Graph — do not treat that as an CLI bug.

## Approvals and workflow automation (Graph)

The CLI now wraps the **delegated** end-user surface — see **`approvals list`** / **`approvals get`** / **`approvals steps`** / **`approvals respond`** (beta `/me/approvals`, scope **`ApprovalSolution.ReadWrite`** — canonical delegated, identifier `6768d3af-4562-48ff-82d2-c5e19eb21b9c`). A narrower **`ApprovalSolutionResponse.ReadWrite`** can be used for read-and-respond only.

Other workflow surfaces (PIM, access-package approvals, identity governance) live under **`/identityGovernance/...`** and **`/roleManagement/...`**. Those paths and permissions change—verify against [Microsoft Graph API reference](https://learn.microsoft.com/en-us/graph/api/overview) before calling production tenants.

Pattern (use only after confirming docs for your scenario):

```bash
m365-agent-cli graph invoke --token "$TOKEN" --beta -X GET "/identityGovernance/entitlementManagement/accessPackageAssignmentApprovals?\$expand=steps&\$top=5"
```

Prefer **Power Automate** / **Approvals** in-product UX for policy-heavy workflows; use **`graph invoke`** only when your tenant has approved the exact API and scopes.

## Teams activity feed — user-targeted (often app-only)

`teams activity-notify` wraps **`POST /me/teamwork/sendActivityNotification`**, **`POST /chats/{id}/sendActivityNotification`**, and **`POST /users/{id}/teamwork/sendActivityNotification`** (use **`--user-id`** for the last; scope / permission **`TeamsActivity.Send`** — **application** consent is typical for the `/users/{id}/…` path).

```bash
m365-agent-cli teams activity-notify --user-id "$USER_ID" --token "$APP_TOKEN" --json-file ./notify.json
```

Equivalent raw invoke:

```bash
m365-agent-cli graph invoke --token "$APP_TOKEN" -X POST "/users/$USER_ID/teamwork/sendActivityNotification" --json-file ./notify.json
```

## Word / PowerPoint on drive items (beta experiments)

The first-class CLI exposes **`word` / `powerpoint`** **preview**, **meta**, **download**, **thumbnails**, and **mirrored** per-item verbs aligned with **`files`** (**upload**, **share**, **permissions**, **versions**, **checkout**, **checkin**, **convert**, **activities**, …). **Excel** on a drive item includes worksheets, ranges, tables (incl. columns and row patch/delete), pivot tables (incl. refresh), names, charts, **workbook-get**, **application-calculate**, sessions (**create** / **refresh** / **close**), and threaded comments under **`excel comments-*`** (Graph **beta**). Workbook **images**, **shapes**, and long **`range()`** method chains remain **`graph invoke`**.

**OpenAPI spike (local msgraph index):** there is **no** stable first-class path analogous to Excel **`…/workbook/comments`** for **Word** or **PowerPoint** drive-hosted document comments. Do **not** expect **`word comments-*`** until Microsoft documents a supported delegated API.

For **Word** or **PowerPoint**, Microsoft may still expose **beta** item facets (e.g. information protection) under `…/drive/items/{id}…` — paths and permissions change; verify the current Graph reference before relying on them.

Illustrative patterns (adjust ids, use delegated token from **`m365-agent-cli login`**, add **`--beta`** when calling beta):

```bash
# List children of the signed-in user's drive root (sanity check)
m365-agent-cli graph invoke -X GET "/me/drive/root/children?\$top=5"

# Beta: always confirm the path in Microsoft Graph docs for your tenant/version
m365-agent-cli graph invoke --beta -X GET "/me/drive/items/{driveItem-id}"

# Example only — extraction / sensitivity APIs vary by license and Graph version; confirm docs
# m365-agent-cli graph invoke --beta -X POST "/me/drive/items/{driveItem-id}/extractSensitivityLabels"
```

For **agent-friendly** editing without unsupported Graph write APIs, prefer **`word download`** / **`powerpoint download`** → local edit → **`files upload`** (see **[`docs/AGENT_WORKFLOWS.md`](./AGENT_WORKFLOWS.md)** § Word / PowerPoint).

**Maintenance:** periodically run **[`scripts/graph-powerpoint-openapi-watch.mjs`](../scripts/graph-powerpoint-openapi-watch.mjs)** (or the `openapi-search` commands in **[`docs/GRAPH_API_GAPS.md`](./GRAPH_API_GAPS.md)** — *PowerPoint Graph API watchlist*) so new **`…/presentation/…`** drive-item paths are noticed early.

**First-party CLI (drive item):** **`files`** (and **`word`/`powerpoint`** mirrors) cover **`list-item`**, **`follow`/`unfollow`**, **`sensitivity-assign`/`sensitivity-extract`**, **`retention-label`/`retention-label-remove`**, and **`permanent-delete`** where Graph exposes them — use **`graph invoke`** only for tenant-specific or undocumented variants.

## Microsoft Viva / employee experience

Prefer **`m365-agent-cli viva`** (Graph **beta**) for the wrapped surfaces: tenant **`/employeeExperience`** (communities, goals, learning providers/contents/activities, engagement roles + **member user** navigation, community **owners** including **UPN** alternate key, **mailboxSettings** / **serviceProvisioningErrors** on owners and role members, async operations), per-user **`employeeExperience`**, **work time** / **item insights**, **admin** and **organization** **`itemInsights`**, **`workHoursAndLocations`**, and **Viva Engage** **meeting** Q&A (**`onlineMeetingConversations`**: conversations, messages **create/patch/delete**, **replies**, **reactions**, **replyTo** / **conversation** / **onlineMeeting** navigation). Scopes and tenant availability: **`docs/GRAPH_SCOPES.md`**.

Use **`graph invoke --beta`** only for **employeeExperience** or meeting-conversation paths the CLI does not yet name (for example rare OData **`$expand`** combinations or preview facets not modeled as subcommands).

## Planner, Microsoft To Do, and Outlook contacts — code-verified vs Graph (beta gaps)

This section is **not** sourced from the parity matrix. It reflects a **static review** of `src/lib/planner-client.ts`, `src/commands/planner.ts`, `src/lib/todo-client.ts`, `src/commands/todo.ts`, and contact APIs in `src/lib/outlook-graph-client.ts` / `src/commands/contacts.ts`, cross-checked against the **local Microsoft Graph OpenAPI index** (`msgraph openapi-search`).

### Planner

| Area | What the CLI uses | Notes |
| --- | --- | --- |
| Core CRUD (plans, buckets, tasks, plan/task **details** PATCH) | **v1.0** (`callGraph` → default `GRAPH_BASE_URL`) | Canonical paths such as `/planner/plans`, `/planner/tasks`, `/planner/buckets`, `/planner/plans/{id}/details`. |
| **Beta-only in this repo** | **`GRAPH_BETA_URL`** | Archive / unarchive plan; `GET …/planner` (**get-me**); **`GET …/planner/myDayTasks`** (**list-my-day-tasks**); **`GET …/planner/recentPlans`** (**list-recent-plans**); **`PATCH …/planner`** (**update-me** with **`--etag`** + **`--json-file`**); favorite + roster plan lists; **`/me/planner/all/delta`**; roster create/get/members; create plan **in roster**; **personal / user container** plan create (**`planner create-plan --me`** → **`POST /me/planner/plans`**); **moveToContainer**; **getUsageRights** (see `planner-client.ts`). Several of these support optional **`--user`** (delegated **`/users/{id}/...`**; may **403**). |
| **Wrapped plan create** | v1 **group** (`POST /planner/plans` + group `container.url`), beta **roster** (`createPlannerPlanInRoster`), or beta **user** container (**`createPlannerPlanForSignedInUser`**, **`planner create-plan --me`**) | Resolves **`GET /me?$select=id`** then posts a **`plannerPlan`** whose **`container`** targets **`…/beta/users/{id}`** with **`type: user`** (per **plannerPlanContainer**). |
| **Delete “details” sub-resources** | **Implemented** | **`planner delete-plan-details`** / **`planner delete-task-details`** (destructive; **`--confirm`**). Alternate URL shapes (group/team-scoped) → **`graph invoke`** if Graph routes you there. |
| Alternate URL shapes | Partially covered | **`moveToContainer`** / **`getUsageRights()`** are implemented on **`/planner/plans/{id}/...`** (beta). Graph also lists **`/me/planner/plans/{id}/...`** variants — if the service returns routing or permission errors, retry the **`/me/...`** form via **`graph invoke --beta`**. |

### Microsoft To Do

| Area | What the CLI uses | Notes |
| --- | --- | --- |
| All `todo-client` REST calls | **v1.0** only | No `GRAPH_BETA_URL` / `--beta` branch in `todo-client.ts`. |
| Task attachments (large upload) | **createUploadSession** on `…/attachments/createUploadSession` | This is the supported large-file path in the CLI. |
| **attachmentSessions** collection | **Wrapped** | **`todo attachment-session`** (list/get/patch/delete + **content** GET/PUT/DELETE). v1 has **no POST** on the collection; sessions typically follow **`createUploadSession`** / **`todo upload-attachment-large`**. |
| **PATCH …/todo** / **DELETE …/todo** | **Wrapped** | **`todo root patch`** / **`todo root delete`** (**`--confirm`** required on delete; optional **`--if-match`**). Still unusual—confirm impact before delete. |

### Outlook contacts (Graph)

| Area | What the CLI uses | Notes |
| --- | --- | --- |
| Folders, contacts, delta, photo, attachments, **open extensions** | **v1.0** | Default extensions: **`/me/contacts/{id}/extensions`** (or delegated **`/users/{id}/...`**). |
| Nested **contactFolders** extension paths | **Wrapped** | **`contacts extension`** accepts **`-f/--folder`** and optional **`--child-folder`** for **`…/contactFolders/{id}/contacts/{contactId}/extensions`** and **`…/childFolders/{id}/…`** shapes. |
| **contactMergeSuggestions** (user settings) | **Implemented** | **`contacts merge-suggestions`** `get` / `set` / `delete` — Graph **beta**; **`--user`** for **`/users/{id}/settings/contactMergeSuggestions`**. |

### Example `graph invoke` patterns (verify paths in current Graph docs)

```bash
# Preferred: personal (user-container) plan — wrapped as:
#   m365-agent-cli planner create-plan --me -t "My plan"

# Beta: same create via raw invoke (only if you need a non-default body)
m365-agent-cli graph invoke --beta -X POST "/me/planner/plans" --json-file ./new-personal-plan.json

# Beta: same move/getUsageRights as first-class commands, but via /me/... if needed
m365-agent-cli graph invoke --beta -X POST "/me/planner/plans/{plan-id}/moveToContainer" \
  --header "If-Match: W/\"etag\"" --json-file ./move-plan.json
m365-agent-cli graph invoke --beta -X GET "/me/planner/plans/{plan-id}/getUsageRights()"
```

## Cross-links

- Prefer **`docs/GRAPH_PRODUCT_PARITY_MATRIX.md`** for workload-level status.
- Prefer **`docs/GRAPH_API_GAPS.md`** for endpoint-level tracking and **closure targets**.
- Assistant-oriented delegation flags: **[`docs/PERSONAL_ASSISTANT_DELEGATION.md`](./PERSONAL_ASSISTANT_DELEGATION.md)**.
