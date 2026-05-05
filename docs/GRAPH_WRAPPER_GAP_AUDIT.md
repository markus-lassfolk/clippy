# Graph wrapper gap audit (discovery + hardening)

**Generated:** 2026-05-05 (maintenance pass; regenerate inventory when Graph call sites change).

**Purpose:** Single place for prioritized **Gap** / **Partial** items, **beta** coverage, **OpenAPI compliance** status, **raw `fetch` bypasses**, **pagination** posture, and **delegation** notes—aligned with [`GRAPH_PRODUCT_PARITY_MATRIX.md`](./GRAPH_PRODUCT_PARITY_MATRIX.md), [`GRAPH_API_GAPS.md`](./GRAPH_API_GAPS.md), and [`GRAPH_INVOKE_BOUNDARIES.md`](./GRAPH_INVOKE_BOUNDARIES.md).

**Path inventory:** [`GRAPH_PATH_INVENTORY.json`](./GRAPH_PATH_INVENTORY.json) — **`npm run graph:inventory`**. Treat as ground truth for implemented URL templates.

**Strict OpenAPI gate:** `npm run verify:graph-compliance` = inventory check + **`GRAPH_OPENAPI_STRICT=1`** against the local [msgraph](https://github.com/merill/graph-skills) OpenAPI index. Unmatched patterns belong in [`scripts/graph-openapi-allowlist.json`](../scripts/graph-openapi-allowlist.json) with a short rationale.

---

## 1. Prioritized gap / partial backlog

| Priority | Workload | Status | API (v1 / beta) | Delegation | Escape hatch | Notes / proposed follow-up |
| --- | --- | --- | --- | --- | --- | --- |
| — | **Teams Phone / PSTN** | **Out of scope** | n/a | n/a | Teams admin / carrier | **Not** a target for this CLI — [`GRAPH_INVOKE_BOUNDARIES.md`](./GRAPH_INVOKE_BOUNDARIES.md). |
| P0 (policy) | **Other Cloud Communications** (Graph call control, call records, …) | Gap | varies | Often app-only | `graph invoke` | High compliance; PSTN product scenarios remain **out of scope** (above). |
| P0 (policy) | Teams RSC / tenant admin | Gap | varies | Often app-only | Admin center / PowerShell / `graph invoke` | By design for delegated CLI. |
| P1 | Word / PowerPoint in-file comments & deck OM | Gap | beta if documented | Delegated where Graph allows | OOXML round-trip; `graph invoke --beta` if Microsoft publishes paths | No stable drive-item path analogous to Excel `workbook/comments`; see [`GRAPH_API_GAPS.md`](./GRAPH_API_GAPS.md) matrices. |
| P1 | Excel images / shapes / deep `range()` chains | Partial | v1 | Delegated | `graph invoke` | Workbook OM long tail. |
| — | Teams activity **`POST /users/{id}/teamwork/sendActivityNotification`** | **Implemented** | v1 | App token typical | `teams activity-notify --user-id` (+ `--token`) | Delegated **`/me/…`** + chat unchanged; see [`GRAPH_INVOKE_BOUNDARIES.md`](./GRAPH_INVOKE_BOUNDARIES.md). |
| P2 | Bookings **staff-availability** | Partial | v1 | **App-only** per Microsoft | `bookings staff-availability --token` | Delegated may 403 by design. |
| — | To Do **attachmentSessions** | **Implemented** | v1 | Delegated | `todo attachment-session` | Distinct from `createUploadSession` (no POST on collection in v1). |
| — | To Do **PATCH/DELETE …/todo** | **Implemented** | v1 | Destructive | `todo root patch` / `todo root delete --confirm` | Unusual; delete is gated. |
| — | Contacts nested folder **extensions** paths | **Implemented** | v1 | Delegated | `contacts extension` with `-f` / `--child-folder` | Default `…/contacts/{id}/extensions` unchanged. |
| P3 | Identity governance / PIM approvals | Gap | beta | varies | `graph invoke --beta` | Out of product scope for first-class CLI; see [`GRAPH_INVOKE_BOUNDARIES.md`](./GRAPH_INVOKE_BOUNDARIES.md). |
| P3 | Directory admin CRUD | Gap | v1 | Admin | `graph invoke` | Delegated `find` / `org` / `people` / `rooms` are wrapped. |
| — | Raw REST / batch | Partial | any | any | `graph invoke`, `graph batch` | Intentional escape hatch. |

**Recently closed (this pass):** **`todo attachment-session`** (list/get/patch/delete + session **content** GET/PUT/DELETE); **`todo root`** get/patch/delete (`--confirm` on delete); **`contacts extension`** folder-scoped paths (`-f/--folder`, `--child-folder`); **`approvals cancel`** (`DELETE /me/approvals/{id}`); **`contacts merge-suggestions`** get/set/delete (beta); **`planner delete-plan-details`** / **`planner delete-task-details`**; OneNote multipart **POST/PATCH** now use **`callGraphAt`** (throttle/retry alignment); **`teams activity-notify --user-id`** (`POST /users/{id}/teamwork/sendActivityNotification`).

---

## 2. Beta and preview modules (code inventory)

| Area | Mechanism | Source files (representative) |
| --- | --- | --- |
| Global beta root | `GRAPH_BETA_URL` / CLI `--beta` | [`graph-constants.ts`](../src/lib/graph-constants.ts), [`drive-location.ts`](../src/lib/drive-location.ts), `files` / `sharepoint` / `site-pages` |
| Approvals | Always beta host | [`graph-approvals-client.ts`](../src/lib/graph-approvals-client.ts) |
| Contact merge suggestions | Always beta host | [`graph-contact-merge-suggestions-client.ts`](../src/lib/graph-contact-merge-suggestions-client.ts) |
| Planner | Mixed v1 + beta | [`planner-client.ts`](../src/lib/planner-client.ts) |
| Viva / employee experience | Beta | [`graph-viva-client.ts`](../src/lib/graph-viva-client.ts), [`graph-viva-tenant-client.ts`](../src/lib/graph-viva-tenant-client.ts), [`graph-viva-meeting-engage-deep.ts`](../src/lib/graph-viva-meeting-engage-deep.ts) |
| Excel workbook comments | Beta | [`graph-excel-comments-client.ts`](../src/lib/graph-excel-comments-client.ts) |
| Copilot | Optional `--beta`; zip/download uses beta URL | [`copilot-graph-client.ts`](../src/lib/copilot-graph-client.ts), [`copilot.ts`](../src/commands/copilot.ts) |
| Teams (subset) | Beta when feature-flagged paths | [`graph-teams-client.ts`](../src/lib/graph-teams-client.ts) |

**Ongoing:** Re-run msgraph `openapi-search` watch scripts from [`GRAPH_API_GAPS.md`](./GRAPH_API_GAPS.md) (Word/PowerPoint, PowerPoint deck) when Microsoft ships new preview APIs.

---

## 3. OpenAPI compliance (strict)

- **Command:** `npm run verify:graph-compliance` (or `npm run verify:graph-openapi:strict`).
- **Allowlist:** Highly templated path patterns are documented in [`scripts/graph-openapi-allowlist.json`](../scripts/graph-openapi-allowlist.json).

---

## 4. Raw `fetch` bypass vs `callGraph` / `callGraphAt`

| Location | Role | Classification |
| --- | --- | --- |
| [`graph-client.ts`](../src/lib/graph-client.ts) | Core Graph JSON, uploads, redirects | **Canonical** — retries, `Retry-After`, 401 refresh hook. **FormData** bodies skip forced `Content-Type: application/json` (multipart). |
| [`graph-advanced-client.ts`](../src/lib/graph-advanced-client.ts) | `graph invoke` / batch | Uses `callGraphAt` → same behavior. |
| [`onenote-graph-client.ts`](../src/lib/onenote-graph-client.ts) | Multipart page create/patch | **Uses `callGraphAt`** for Graph JSON host + retry stack (multipart). |
| [`copilot-graph-client.ts`](../src/lib/copilot-graph-client.ts) | Package zip GET/PUT | **Exception** — binary/stream URLs on beta host. |
| [`todo-client.ts`](../src/lib/todo-client.ts) | Chunk upload to `uploadUrl` | **Exception** — session URL from Graph. |
| [`graph-attachment-upload-session.ts`](../src/lib/graph-attachment-upload-session.ts) | PUT to session URL | **Exception** — same as upload sessions. |
| [`graph-meeting-recordings-client.ts`](../src/lib/graph-meeting-recordings-client.ts) | `redirect: manual` for content | **Exception** — recording/transcript download flow. |
| [`graph-client.ts`](../src/lib/graph-client.ts) | Async job monitor `fetch` | **Exception** — monitor URL validation encloses allowed hosts. |
| OAuth / EWS / npm | Non-Graph | Out of scope. |

---

## 5. Pagination and truncation

| Pattern | Examples | Notes |
| --- | --- | --- |
| `fetchAllPages` + `GRAPH_PAGE_DELAY_MS` | `files`, `outlook-graph` with `--all`, `contacts`, calendar lists, `places`, `sharepoint`, `excel comments`, `site-pages`, **`approvals list --all`** | Preferred for OData collections. |
| Explicit `--next` / `nextLink` printing | `meeting recordings-all`, deltas, **`approvals list --next`** | Script-friendly continuation. |
| Delta + `--state-file` | `todo`, `planner`, `contacts`, mail/calendar deltas, recordings/transcripts delta | Durable sync. |
| `--all-pages` | `sharepoint items` | Named full walk. |

---

## 6. Delegation (`--user` / shared access)

Canonical narrative: [`PERSONAL_ASSISTANT_DELEGATION.md`](./PERSONAL_ASSISTANT_DELEGATION.md) (§7).

---

## 7. Capability matrix / scopes

- Source: [`graph-capability-matrix.ts`](../src/lib/graph-capability-matrix.ts) → **`verify-token --capabilities`**.
- **Regenerate doc:** `npm run docs:graph-permission-matrix` after matrix edits; **`npm run docs:graph-permission-matrix:check`** in CI.

---

## 8. Maintenance commands (checklist)

```bash
npm run graph:inventory
npm run verify:graph-compliance
npm run docs:graph-permission-matrix:check
```

Optional: `bash ~/.cursor/skills/msgraph/scripts/run.sh openapi-search --query "<workload>" --limit 25` for drift against Microsoft’s OpenAPI.
