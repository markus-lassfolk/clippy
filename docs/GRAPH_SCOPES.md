# Microsoft Graph OAuth scopes (m365-agent-cli)

This document lists **delegated** permissions the CLI is designed to use. **Source of truth in code:** [`src/lib/graph-oauth-scopes.ts`](../src/lib/graph-oauth-scopes.ts) (`GRAPH_DEVICE_CODE_LOGIN_SCOPES` for `login`, `GRAPH_REFRESH_SCOPE_CANDIDATES` for token refresh in [`graph-auth`](../src/lib/graph-auth.ts)).

Configure the same permissions on your **Entra ID app registration** (API permissions → Microsoft Graph → Delegated). Then run **`m365-agent-cli login`** so the refresh token includes them. Use **`m365-agent-cli verify-token`** to inspect granted `scp` claims. Use **`m365-agent-cli verify-token --capabilities`** for a read/write checklist of CLI feature areas (Planner, SharePoint, mail, Teams, …) inferred from permission names on the token; add **`--json`** for machine-readable output.

**Office 365 Exchange Online:** add **`EWS.AccessAsUser.All`** (delegated) for EWS-backed commands when `M365_EXCHANGE_BACKEND` is `ews` or `auto` (see [`ENTRA_SETUP.md`](./ENTRA_SETUP.md)).

---

## Full scope set (recommended)

| Scope | Purpose in this CLI |
| --- | --- |
| `offline_access` | Refresh tokens |
| `User.Read` | Sign-in profile; **`/me`**; **`org manager`** / **`org direct-reports`** without **`--user`** (**GET /me/manager**, **GET /me/directReports**) where tenant policy allows |
| `Calendars.ReadWrite` | Own calendar read/write |
| `Calendars.Read.Shared` | Delegated / shared calendars (`/users/{upn}/calendar/...`) |
| `Calendars.ReadWrite.Shared` | Same, with write |
| `Mail.Send` | **`POST /me/sendMail`** and sending mail via Graph (explicit; use with `Mail.ReadWrite` for full mail UX) |
| `Mail.ReadWrite` | Own mailbox mail APIs |
| `Mail.Read.Shared` | Mail in mailboxes the user can access (delegated/shared) |
| `Mail.ReadWrite.Shared` | Same, with write (where applicable) |
| `MailboxSettings.ReadWrite` | Mailbox settings, OOF, categories, rules-related settings |
| `Place.Read.All` | Places API — `rooms`, room resolution in `create-event` / Places |
| `People.Read` | `GET /me/people` — `find` (people/relevant contacts) |
| `User.Read.All` | `GET /users` directory search — `find` (user query); **`org manager`** / **`org direct-reports`** with **`--user`** (**GET /users/{id}/manager**, **GET /users/{id}/directReports**); **often requires admin consent** |
| `Files.ReadWrite.All` | OneDrive / **`files`** (incl. **`delta`**, **`shared-with-me`**, **`thumbnails`**, copy/move, **`invite`**, **`permissions`**, **`permission-remove`**, **`permission-update`**, share links, upload/delete) — any drive root selected by flags; **`word`**/**`powerpoint`** **preview** / **meta** / **download** / **thumbnails**; **`excel`** workbook — worksheets **get/add/update/delete**; **range** read + **range-patch** + **range-clear**; **used-range**; **tables** (list/get/create/patch/delete, rows list/add/patch/delete, columns list/get/patch); **pivot-tables** + pivot CRUD/refresh; **names** list + **name-get** + **worksheet-names** / **worksheet-name-get**; **charts**; **workbook-get**; **application-calculate**; **session-create** / **session-refresh** / **session-close**; **`comments-*`** (Graph **beta** workbook comments); optional **`--session-id`** on mutating **`excel`** calls |
| `Sites.ReadWrite.All` | **`sharepoint`** lists/items/**`resolve-site`**/**`items-delta`**/**`get-item`**/**`delete-item`**; **`site-pages`** |
| `Tasks.ReadWrite` | Microsoft To Do |
| `Group.ReadWrite.All` | Planner (groups), group-related Graph calls; **`teams members`**, **`teams channel-members`**, **`teams team-member-add`**, **`teams channel-member-add`**, **`teams apps`**, **`teams tabs`** / **tab-*** (narrower delegated scopes exist for read-only vs tab/member writes — see Graph permission reference) |
| `Contacts.ReadWrite` | **`contacts`** — `/me/contacts`, `/me/contactFolders`, photo, file + **reference (link)** attachments, delta, `$search`, `$filter` on list |
| `Contacts.Read.Shared` | Read contacts in **shared / delegated** mailboxes (`--user` on `contacts`) |
| `Contacts.ReadWrite.Shared` | Create/update/delete contacts for mailboxes you have delegate access to |
| `OnlineMeetings.ReadWrite` | **`meeting`** — `POST/PATCH/DELETE/GET /me/onlineMeetings` (standalone Teams meeting; **`meeting create --json-file`** for full Graph body). **`Calendars.ReadWrite`** + **`create-event … --teams`** — calendar invitations with Teams; parse **`--json`** → `event.teamsMeeting` / `event.onlineMeeting` for assistants. |
| `Notes.ReadWrite.All` | **`onenote`** — notebooks / section groups / sections (CRUD), pages (list, get, HTML, export, create, **delete**, **patch-page-content**, **copy-page**), async **operation** poll for copy |
| `Team.ReadBasic.All` | **`teams`** — joined teams, team metadata (`GET /me/joinedTeams`, `GET /users/{id}/joinedTeams` when using **`teams list --user`**, `GET /teams/{id}`) |
| `Channel.ReadBasic.All` | **`teams channels`**, **`teams all-channels`**, **`teams incoming-channels`**, **`teams primary-channel`**, **`teams channel-get`**, **`teams channel-files-folder`** — list/get channel (`/channels`, `/allChannels`, `/incomingChannels`, `primaryChannel`, `channels/{id}`, `channels/{id}/filesFolder`) |
| `ChannelMessage.Read.All` | **`teams messages`**, **`teams channel-message-get`**, **`teams message-replies`** — channel messages and thread replies; **delegated admin consent** often required |
| `ChannelMessage.Send` | **`teams channel-message-send`**, **`teams channel-message-reply`** — `POST …/channels/{id}/messages` and `…/messages/{id}/replies` |
| `Presence.Read.All` | **`presence me`**, **`presence user`**, **`presence bulk`** (`POST /communications/getPresencesByUserId`) |
| `Presence.ReadWrite` | **`presence set-me`**, **`presence set-user`**, **`presence clear-me`**, **`presence clear-user`** — set/clear presence session |
| `Bookings.ReadWrite.All` | **`bookings`** — delegated read/write as listed; **`staff-availability`** is **not** delegated per Microsoft (use **app-only** token) |
| `Chat.ReadWrite` | **`teams chats`**, **`teams chat-get`**, **`teams chat-pinned`**, **`teams chat-messages`**, **`teams chat-message-get`**, **`teams chat-message-replies`**, **`teams chat-members`**, **`teams chat-message-send`**, **`teams chat-message-reply`** |
| `ExternalItem.Read.All` | **`copilot retrieval`** with `dataSource=externalItem` — Copilot connectors content ([Retrieval API](https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/api/ai-services/retrieval/overview)); SharePoint/OneDrive sources need **`Files.Read.All`** + **`Sites.Read.All`** (or **`Files.ReadWrite.All`** + **`Sites.ReadWrite.All`** per Microsoft) |
| `Reports.Read.All` | **`copilot reports`** (usage summary / trend / user detail) — plus an admin “reports reader” role per [Microsoft Graph report authorization](https://learn.microsoft.com/en-us/graph/reportroot-authorization) |
| `CopilotPackages.Read.All` / `CopilotPackages.ReadWrite.All` | **`copilot packages`** list/get vs update/block/unblock/reassign ([package API](https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/api/admin-settings/package/overview)); requires Microsoft Agent 365 licensing per Microsoft |
| `OnlineMeetingAiInsight.Read.All` | **`copilot meeting-insights-list`**, **`meeting-insight-get`** ([Meeting insights](https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/api/ai-services/meeting-insights/callaiinsight-get)) |
| `OnlineMeetingTranscript.Read.All` | **Copilot Chat API** (`copilot chat`, `chat-stream`) — part of the required delegated bundle in [Chat API permissions](https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/api/ai-services/chat/copilotroot-post-conversations#permissions) |
| `People.Read.All` | **Copilot Chat API** bundle; also satisfies **`find`** people-style directory reads where applicable |
| `AiEnterpriseInteraction.Read` | **Delegated** [change notifications](https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/api/ai-services/change-notifications/aiinteraction-changenotifications) for a user’s interactions; **`subscribe copilot-interactions --user …`** |
| `AiEnterpriseInteraction.Read.All` | **Application-only** interaction export and tenant-wide interaction subscriptions (not in typical user-delegated `login` token) |
| `ApprovalSolution.ReadWrite` | **`approvals list`** / **`approvals get`** / **`approvals steps`** / **`approvals respond`** — Teams Approvals app + Power Automate approvals via **`/me/approvals`** (beta). Delegated permission identifier `6768d3af-4562-48ff-82d2-c5e19eb21b9c` (DisplayText: *Read, create, and respond to approvals*; admin consent required). A narrower **`ApprovalSolutionResponse.ReadWrite`** exists for read-and-respond only. |
| `OnlineMeetingRecording.Read.All` | **`meeting recordings`** / **`recording-download`** / **`recordings-all`** (per-meeting + tenant-wide `getAllRecordings(...)` and `recordings/delta()`); requires tenant Stream/Teams policy. |
| `TeamsActivity.Send` | **`teams activity-notify`** — **delegated** **`POST /me/teamwork/sendActivityNotification`** and **`POST /chats/{id}/sendActivityNotification`**; the app-only **`/users/{id}/teamwork/sendActivityNotification`** path is intentionally not wrapped (use **`graph invoke`** with an app token). |
| *(entity-specific)* | **`graph-search`** — Microsoft Graph Search (`POST /search/query`) uses the least-privilege permission for each entity type (e.g. mail → Mail.Read, files → Files.Read.All); see Graph Search API docs |
| — | **`graph invoke`** / **`graph batch`** — arbitrary JSON Graph paths and `$batch` (see command help); use for APIs not wrapped as dedicated subcommands |

**Note:** `Group.ReadWrite.All` implies broad group read/write. For **`find`** group listing, this is sufficient; a narrower `Group.Read.All` is **not** requested separately to avoid redundant consent alongside `Group.ReadWrite.All`.

---

## Admin consent

These commonly require **admin consent** in tenant consent policies (especially **`User.Read.All`**, **`Place.Read.All`**). If refresh fails after login, check the Entra **Enterprise applications** → your app → **Permissions** and use **Grant admin consent**, or ask an admin to approve.

---

## Refresh fallback behavior

[`graph-auth`](../src/lib/graph-auth.ts) tries several scope strings when refreshing. It includes a candidate **without** `User.Read.All` so users who cannot obtain admin consent for directory read may still refresh tokens for mail/calendar/files-heavy operations.

---

## Related docs

- [`ENTRA_SETUP.md`](./ENTRA_SETUP.md) — portal steps and automated scripts  
- [README](../README.md) — authentication overview  
- [`MIGRATION_TRACKING.md`](./MIGRATION_TRACKING.md) — Graph vs EWS, `--mailbox` behavior  
- [`GRAPH_TROUBLESHOOTING.md`](./GRAPH_TROUBLESHOOTING.md) — OData headers, `$search`, `/me/people`, and consistency quirks  

**Maintainers:** To sanity-check live endpoints, `$filter`, and permission requirements against current Microsoft Graph docs, use the **msgraph** Cursor skill ([graph.pm introduction](https://graph.pm/getting-started/introduction/)); install with `npx skills add merill/msgraph` if you use Cursor skills.

*Last updated: 2026-05-05 — **`excel`** expanded verbs documented under **`Files.ReadWrite.All`** (same scope as before). Phase 1–3 scopes: **`ApprovalSolution.ReadWrite`** (**`approvals`**, verified canonical delegated name; identifier `6768d3af-…`), **`OnlineMeetingRecording.Read.All`** (**`meeting recordings*`**), **`TeamsActivity.Send`** (**`teams activity-notify`**); aligned with `graph-oauth-scopes.ts`. **`insights`**, **`files recent`/`activities`/`preview`**, **`sharepoint followed-sites`/`follow`/`unfollow`**, **`groups`**.*

