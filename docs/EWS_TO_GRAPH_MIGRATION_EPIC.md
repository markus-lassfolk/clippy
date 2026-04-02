# Epic: Migrate Exchange Web Services (EWS) to Microsoft Graph

**Status:** Planning / in progress  
**GitHub Epic:** [#204 — EWS → Microsoft Graph migration](https://github.com/markus-lassfolk/m365-agent-cli/issues/204) (17 sub-issues linked under the epic in GitHub)  
**Driver:** [Exchange Online retirement of EWS](https://learn.microsoft.com/en-us/graph/migrate-exchange-web-services-overview) (phased; confirm dates in Microsoft docs and Message Center).  
**Strategy:** Phased migration with **Microsoft Graph as the primary implementation** and **EWS as fallback** until each slice is verified; then remove EWS for that slice.

---

## How to track this in GitHub

1. Use Epic **[#204](https://github.com/markus-lassfolk/m365-agent-cli/issues/204)** and its **sub-issues** (see inventory table below).
2. New slices: **New issue** → template **EWS → Graph migration task** (`.github/ISSUE_TEMPLATE/ews-graph-migration.yml`), then link as a sub-issue of #204 if needed (`gh api graphql` `addSubIssue`, or GitHub UI).
3. **One-off script** that created #204–#221: `scripts/create-ews-graph-issues.ps1` — **do not re-run** (it would create duplicate issues). Use it only as a reference for `gh` + GraphQL `addSubIssue`.

Labels used: `epic`, `migration`, `ews`, `graph`.

---

## Fallback model (implementation pattern)

Until a slice is marked **EWS removed**, implementations should follow a consistent pattern:

1. **Single entry point per domain** (e.g. `calendar-read`, `mail-send`) that chooses backend:
   - `graph` — use Graph only (for tenants already cut off from EWS or for tests).
   - `ews` — use EWS only (legacy / emergency).
   - `auto` *(default during migration)* — try Graph first; on **definitive** failure (e.g. known unsupported case, or opt-in retry policy), fall back to EWS once.
2. **Configuration** (names are proposals; implement in code when starting Phase 1):
   - Env: `M365_MAIL_BACKEND`, `M365_CALENDAR_BACKEND`, etc., with values `graph` | `ews` | `auto`, **or** one `M365_EXCHANGE_BACKEND=auto`.
   - Document in README when introduced.
3. **Observability:** Log which backend served each request (debug/`--verbose` only) so support can see Graph vs EWS.
4. **Tests:** For each migrated command, add tests for Graph path; keep EWS mocks until EWS deletion phase.

**Definition of “slice complete”:** Graph path is default in `auto`, feature parity documented, EWS path behind flag for that slice only until hard cutover.

---

## Reference

- [Migrate EWS apps to Microsoft Graph (overview)](https://learn.microsoft.com/en-us/graph/migrate-exchange-web-services-overview)
- [EWS to Graph API mapping](https://learn.microsoft.com/en-us/graph/migrate-exchange-web-services-api-mapping)
- Repo architecture note: `docs/ARCHITECTURE.md` (EWS vs Graph priority)

---

## Inventory: EWS touchpoints in this repo

| Area | Commands / modules | Graph direction | Notes | Issue | Status |
| --- | --- | --- | --- | --- | --- |
| Phase 0 foundation | Router, env, Azure scopes inventory | — | Backend `auto`, permissions audit | [#205](https://github.com/markus-lassfolk/m365-agent-cli/issues/205) | ⬜ |
| Calendar read | `calendar` | `GET calendarView` / shared calendars | Replace `getCalendarEvents` | [#206](https://github.com/markus-lassfolk/m365-agent-cli/issues/206) | ⬜ |
| Free-busy / findtime | `findtime`, parts of schedule | `calendar/getSchedule` | Drop `getScheduleViaOutlook` | [#207](https://github.com/markus-lassfolk/m365-agent-cli/issues/207) | ⬜ |
| Whoami | `whoami` | `/me` (+ optional mailboxSettings) | Drop `getOwaUserInfo` / ResolveNames | [#208](https://github.com/markus-lassfolk/m365-agent-cli/issues/208) | ⬜ |
| Mail CRUD + actions | `mail` | Messages, move, patch, send | Large attachment / MIME edge cases | [#209](https://github.com/markus-lassfolk/m365-agent-cli/issues/209) | ⬜ |
| Send | `send` | `sendMail` / draft send | | [#210](https://github.com/markus-lassfolk/m365-agent-cli/issues/210) | ⬜ |
| Drafts | `drafts` | Graph draft messages | | [#211](https://github.com/markus-lassfolk/m365-agent-cli/issues/211) | ⬜ |
| Folders | `folders` | mailFolders | | [#212](https://github.com/markus-lassfolk/m365-agent-cli/issues/212) | ⬜ |
| Todo link | `todo --link` | `getEmail` → Graph get message | Small change | [#213](https://github.com/markus-lassfolk/m365-agent-cli/issues/213) | ⬜ |
| Calendar write | `create-event`, `update-event`, `delete-event` | Events API + online meetings | Time zones, recurrence, Teams | [#214](https://github.com/markus-lassfolk/m365-agent-cli/issues/214) | ⬜ |
| Meeting response | `respond` | Accept/decline/tentative via Graph | Shared mailbox = `/users/{id}/` | [#215](https://github.com/markus-lassfolk/m365-agent-cli/issues/215) | ⬜ |
| Forward / counter | `forward-event`, `counter` | Event forward / propose times | Verify Graph equivalents | [#216](https://github.com/markus-lassfolk/m365-agent-cli/issues/216) | ⬜ |
| Auto-reply (EWS) | `auto-reply` | Deprecate in favor of Graph `oof` / mailboxSettings | Align with existing `oof` command | [#217](https://github.com/markus-lassfolk/m365-agent-cli/issues/217) | ⬜ |
| Delegates | `delegates`, `delegate-client.ts` | Calendar permission / share APIs | **No 1:1 EWS delegate matrix** — product redesign | [#218](https://github.com/markus-lassfolk/m365-agent-cli/issues/218) | ⬜ |
| Auth | `auth.ts`, env `EWS_*` | Single token + Graph scopes | Align with single-cache epic in `docs/GOALS.md` | [#219](https://github.com/markus-lassfolk/m365-agent-cli/issues/219) | ⬜ |
| Tests / mocks | `src/test/mocks`, integration tests | Graph-shaped mocks | | [#220](https://github.com/markus-lassfolk/m365-agent-cli/issues/220) | ⬜ |
| Docs | README, ENTRA_SETUP, SKILL | Remove EWS setup when cut over | | [#221](https://github.com/markus-lassfolk/m365-agent-cli/issues/221) | ⬜ |

Legend: ⬜ not started · 🟡 in progress · ✅ done (EWS fallback removable for that row)

---

## Phased roadmap

### Phase 0 — Foundation

- [x] Create GitHub Epic + child issues from inventory table ([#204](https://github.com/markus-lassfolk/m365-agent-cli/issues/204), [#205](https://github.com/markus-lassfolk/m365-agent-cli/issues/205)–[#221](https://github.com/markus-lassfolk/m365-agent-cli/issues/221))  
- [ ] Agree env vars / `auto` fallback behavior (see above)  
- [ ] Add minimal backend router module stub (no behavior change yet) or document “first PR adds router”  
- [ ] Inventory Azure AD app permissions needed for full Graph parity (mail, calendar, mailboxSettings, …)

**Exit:** Epic linked; Phase 1 issue open.

### Phase 1 — Read-only paths

- [ ] `whoami` → Graph  
- [ ] `calendar` list/view → Graph (+ shared/delegated calendar rules)  
- [ ] `findtime` / schedule: remove remaining EWS-only schedule calls  
- [ ] Read paths keep EWS fallback via `auto` until verified  

**Exit:** Default `auto` uses Graph for reads; EWS fallback tested.

### Phase 2 — Mail stack

- [ ] `mail` (list/read/download unchanged pattern; mutations → Graph)  
- [ ] `send`  
- [ ] `drafts`  
- [ ] `folders`  

**Exit:** Mail commands use Graph in `auto`; EWS optional per env.

### Phase 3 — Calendar writes + meeting actions

- [ ] `create-event`, `update-event`, `delete-event`  
- [ ] `respond`  
- [ ] `forward-event`, `counter`  

**Exit:** Calendar lifecycle on Graph in `auto`.

### Phase 4 — Rules / OOF consolidation

- [ ] Ensure inbox rules are Graph-only (`rules` today)  
- [ ] Merge or deprecate `auto-reply` vs `oof`  

**Exit:** No EWS for OOF/rules.

### Phase 5 — Delegates (redesign)

- [ ] Spike: Graph calendar delegate/share flows vs current CLI UX  
- [ ] New subcommands or breaking change doc  
- [ ] Implement; EWS fallback only if still required for gap (document gap)  

**Exit:** Documented parity or known limitations.

### Phase 6 — EWS removal

- [ ] Remove `callEws`, `ews-client` usage, SOAP mocks  
- [ ] Remove `EWS_REFRESH_TOKEN` / separate EWS cache (single Graph auth)  
- [ ] Update Entra scripts, README, skills  

**Exit:** No EWS in repo; CI green.

---

## Child issue checklist (copy into each issue)

- [ ] Scope: one row (or one small group) from the inventory table  
- [ ] Graph implementation + tests  
- [ ] `auto` fallback to EWS (until slice signed off)  
- [ ] README / `--help` if user-visible flags added  
- [ ] This doc updated: Issue #, Status ✅ for that row  

---

## Open decisions (record answers here)

| Question | Decision |
| --- | --- |
| One env var vs per-area (`MAIL`, `CALENDAR`, …)? | *TBD* |
| Default during migration: `auto` everywhere? | *TBD* (recommended: yes) |
| Breaking CLI changes for `delegates`? | *TBD* |

---

*Last updated: 2026-03-31 — GitHub Epic #204 and sub-issues #205–#221 created.*
