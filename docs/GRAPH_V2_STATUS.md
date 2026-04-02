# Graph-first (`dev_v2`) — status

**Branch:** `dev_v2`  
**Epic:** [#204 — EWS → Microsoft Graph migration](https://github.com/markus-lassfolk/m365-agent-cli/issues/204)  
**Goal:** Move toward **Microsoft Graph as the default** for Exchange-related flows, with **`M365_EXCHANGE_BACKEND`** to opt into EWS or `auto` during migration.

This file is the working log for `dev_v2`. Update it when slices land or decisions change.

---

## Configuration

| Env | Values | Default on `dev_v2` |
| --- | --- | --- |
| `M365_EXCHANGE_BACKEND` | `graph` · `ews` · `auto` | **`graph`** (Graph-only for commands that honor the router) |

- **`graph`** — Graph APIs only (`resolveGraphAuth` + Graph REST).  
- **`ews`** — Legacy EWS only (`resolveAuth` + SOAP) where implemented.  
- **`auto`** — Try Graph first, then EWS if Graph auth or the call fails (per command).

Implementation: `src/lib/exchange-backend.ts`.

---

## Done (this branch)

| Item | Notes |
| --- | --- |
| Phase 0 stub | `getExchangeBackend()`, `DEFAULT_EXCHANGE_BACKEND='graph'`, helpers for tests |
| `whoami` | Uses **`GET /me`** on Graph when `graph` or `auto` (Graph path); EWS path when `ews`; `auto` falls back to EWS |
| `folders` | Graph: `listAllMailFoldersRecursive`, create/rename/delete via Graph mail folders; `ews` / `auto` unchanged semantics |
| `send` | Graph: `sendMail` + `buildGraphSendMailPayload` (file attachments); **`--attach-link` not on Graph** — use `ews` or `auto`; `auto` falls back to EWS on Graph failure |
| `mail` | Graph path: **list** + **`--read`** only (`mail-graph.ts`); search/flags/download/reply/etc. → EWS or error on `graph` |
| `drafts` | Graph path: **list drafts** only; create/edit/send/delete/read-by-id still **EWS** |
| `outlook-graph-client` | `listAllMailFoldersRecursive`; `MessagesQueryOptions.skip`; `OutlookMessage` body / lastModified |
| Unit tests | `src/lib/exchange-backend.test.ts` |

---

## Next (priority order — aligns with epic phases)

1. **Calendar** — Default `calendar` / writes via Graph; keep `M365_EXCHANGE_BACKEND=ews` for escape hatch — [#206](https://github.com/markus-lassfolk/m365-agent-cli/issues/206), [#214](https://github.com/markus-lassfolk/m365-agent-cli/issues/214).  
3. **`findtime`** — Graph `getSchedule` / align with `schedule` — [#207](https://github.com/markus-lassfolk/m365-agent-cli/issues/207).  
4. **Auth consolidation** — Single refresh token / cache where possible — [#219](https://github.com/markus-lassfolk/m365-agent-cli/issues/219).  
5. **Delegates / auto-reply** — Graph or deprecate — [#217](https://github.com/markus-lassfolk/m365-agent-cli/issues/217), [#218](https://github.com/markus-lassfolk/m365-agent-cli/issues/218).  
6. **Phase 6** — Remove EWS client usage when parity is verified — epic Phase 6.

---

## Open decisions

| Topic | Status |
| --- | --- |
| Default on `main` after merge: keep `graph` vs switch to `auto` | TBD before merge |
| Per-area env (`M365_MAIL_BACKEND`, …) vs single `M365_EXCHANGE_BACKEND` | Single var for now; split if needed |

---

*Last updated: 2026-04-02 — Mail stack: `folders`, `send`, `mail` (list/read), `drafts` (list) on Graph.*
