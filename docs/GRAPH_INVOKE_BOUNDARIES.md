# Graph API: intentional CLI boundaries (`graph invoke`)

Some Microsoft Graph areas are **not** wrapped as first-class subcommands in `m365-agent-cli`. This document explains why and how to call them safely.

## Cloud Communications (calls, PSTN, meetings control)

**Voice**, **PSTN**, **call records**, and related endpoints carry **high compliance and consent** requirements and vary by tenant licensing. Use **`m365-agent-cli graph invoke`** with paths from [Microsoft Graph documentation](https://learn.microsoft.com/en-us/graph/api/overview) after your organization approves the permissions.

Do **not** expect copy-paste examples to work without tenant policy and admin consent.

> **Carve-out:** **delegated meeting recordings & transcripts** (per-meeting + tenant-wide `getAllRecordings(...)` / `getAllTranscripts(...)` and their `delta` functions) **are** wrapped — see **`meeting recordings*`** / **`meeting transcripts*`** with **`OnlineMeetingRecording.Read.All`** / **`OnlineMeetingTranscript.Read.All`**. Tenant Stream/Teams policies can still return 403; that's policy, not CLI breakage.

### Example `graph invoke` recipes (adjust tenant paths and ids)

Replace `$TOKEN` with a bearer token that has the listed permission (often **application** permission + admin consent). Paths use **v1.0** relative to `https://graph.microsoft.com`.

**List call records (requires CallRecords.Read.All or equivalent)**

```bash
m365-agent-cli graph invoke --token "$TOKEN" GET "/v1.0/communications/callRecords?\$top=5"
```

**Get PSTN call session id / debugging** — follow Microsoft docs for the exact resource; pattern:

```bash
m365-agent-cli graph invoke --token "$TOKEN" GET "/v1.0/communications/<resource-from-docs>"
```

Prefer **Teams admin center** or **tenant-approved** automation for PSTN policy changes.

## Teams resource-specific consent (RSC) / app-only admin scenarios

Many Teams admin and tenant-wide operations require **application** permissions or policies that **delegated** interactive users cannot satisfy. Prefer **Microsoft Teams admin center**, **PowerShell**, or dedicated automation with app-only tokens—not the default delegated **`login`** flow.

For delegated chat/channel scenarios already covered, use **`teams`**. To script against files shown in a channel’s **Files** tab, use **`teams channel-files-folder`** (with team and channel ids) and then **`files list --drive-id … --folder …`** (wrapped); admin-only paths remain **`graph invoke`**.

### Example (application permission — illustrative only)

```bash
m365-agent-cli graph invoke --token "$APP_TOKEN" GET "/v1.0/groups?\$filter=resourceProvisioningOptions/Any(x:x eq 'Team')&\$top=5"
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
m365-agent-cli graph invoke --token "$TOKEN" --beta GET "/beta/identityGovernance/entitlementManagement/accessPackageAssignmentApprovals?\$expand=steps&\$top=5"
```

Prefer **Power Automate** / **Approvals** in-product UX for policy-heavy workflows; use **`graph invoke`** only when your tenant has approved the exact API and scopes.

## Teams activity feed — app-only path

`teams activity-notify` wraps the **delegated** **`POST /me/teamwork/sendActivityNotification`** and **`POST /chats/{id}/sendActivityNotification`** flows (scope **`TeamsActivity.Send`**). The **app-only** **`POST /users/{id}/teamwork/sendActivityNotification`** path requires a different consent flow and is intentionally not wrapped — call it via **`graph invoke`** with an application token:

```bash
m365-agent-cli graph invoke --token "$APP_TOKEN" POST "/v1.0/users/$USER_ID/teamwork/sendActivityNotification" --json-file ./notify.json
```

## Word / PowerPoint on drive items (beta experiments)

The first-class CLI exposes **`word` / `powerpoint`** **preview**, **meta**, **download**, and **thumbnails** (same as **`files thumbnails`**). **Excel** on a drive item includes worksheets, ranges, tables (incl. columns and row patch/delete), pivot tables (incl. refresh), names, charts, **workbook-get**, **application-calculate**, sessions (**create** / **refresh** / **close**), and threaded comments under **`excel comments-*`** (Graph **beta**). Workbook **images**, **shapes**, and long **`range()`** method chains remain **`graph invoke`**.

**OpenAPI spike (local msgraph index):** there is **no** stable first-class path analogous to Excel **`…/workbook/comments`** for **Word** or **PowerPoint** drive-hosted document comments. Do **not** expect **`word comments-*`** until Microsoft documents a supported delegated API.

For **Word** or **PowerPoint**, Microsoft may still expose **beta** item facets (e.g. information protection) under `…/drive/items/{id}…` — paths and permissions change; verify the current Graph reference before relying on them.

Illustrative patterns (adjust ids, use delegated token from **`m365-agent-cli login`**, add **`--beta`** when calling beta):

```bash
# List children of the signed-in user's drive root (sanity check)
m365-agent-cli graph invoke GET "/v1.0/me/drive/root/children?\$top=5"

# Beta: always confirm the path in Microsoft Graph docs for your tenant/version
m365-agent-cli graph invoke --beta GET "/beta/me/drive/items/{driveItem-id}"

# Example only — extraction / sensitivity APIs vary by license and Graph version; confirm docs
# m365-agent-cli graph invoke --beta POST "/beta/me/drive/items/{driveItem-id}/extractSensitivityLabels"
```

For **agent-friendly** editing without unsupported Graph write APIs, prefer **`word download`** / **`powerpoint download`** → local edit → **`files upload`** (see **[`docs/AGENT_WORKFLOWS.md`](./AGENT_WORKFLOWS.md)** § Word / PowerPoint).

**Maintenance:** periodically run **[`scripts/graph-powerpoint-openapi-watch.mjs`](../scripts/graph-powerpoint-openapi-watch.mjs)** (or the `openapi-search` commands in **[`docs/GRAPH_API_GAPS.md`](./GRAPH_API_GAPS.md)** — *PowerPoint Graph API watchlist*) so new **`…/presentation/…`** drive-item paths are noticed early.

## Cross-links

- Prefer **`docs/GRAPH_PRODUCT_PARITY_MATRIX.md`** for workload-level status.
- Prefer **`docs/GRAPH_API_GAPS.md`** for endpoint-level tracking and **closure targets**.
- Assistant-oriented delegation flags: **[`docs/PERSONAL_ASSISTANT_DELEGATION.md`](./PERSONAL_ASSISTANT_DELEGATION.md)**.
