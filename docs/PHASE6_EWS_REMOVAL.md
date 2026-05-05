# Phase 6 — EWS removal (checklist)

**Epic:** [EWS → Microsoft Graph migration](./EWS_TO_GRAPH_MIGRATION_EPIC.md) — GitHub [#204](https://github.com/markus-lassfolk/m365-agent-cli/issues/204).

**Status:** Not executed in code — EWS remains required for **classic delegates** (`delegates add|update|remove`), **`auto-reply`** (inbox-rule templates), and **`auto`** / **`ews`** backend flows until product owners sign off.

## Prerequisites before deleting `ews-client.ts`

1. **Delegates:** Either Microsoft exposes Graph parity for the full folder-level delegate matrix, or this CLI **drops** classic delegate mutations and documents migration to **`delegates calendar-share`** / admin tooling only.
2. **`auto-reply`:** Deprecated and removed or replaced by documented **`oof`** + **`rules`** workflows only.
3. **Tenant coverage:** Default **`M365_EXCHANGE_BACKEND=graph`** validated across core mail/calendar commands with no production reliance on EWS fallback.
4. **Tests:** Replace SOAP mocks and integration paths that assume EWS; CI green with Graph-only defaults.

## When Phase 6 ships (expected code changes)

- Remove **`src/lib/ews-client.ts`**, **`src/lib/delegate-client.ts`** (if EWS-only), and EWS branches from commands that still dual-path.
- Simplify **`docs/ENTRA_SETUP.md`**: drop **Office 365 Exchange Online** `EWS.AccessAsUser.All` where no longer needed.
- Update **`skills/m365-agent-cli/SKILL.md`** and README to Graph-first-only auth.
- Remove legacy env aliases (**`EWS_REFRESH_TOKEN`** / **`GRAPH_REFRESH_TOKEN`**) after a deprecation window if desired.

Until then, keep **`M365_REFRESH_TOKEN`** + unified **`token-cache-*.json`** as documented in [`ARCHITECTURE.md`](./ARCHITECTURE.md).
