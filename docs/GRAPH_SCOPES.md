# Microsoft Graph OAuth scopes (m365-agent-cli)

This document lists **delegated** permissions the CLI is designed to use. **Source of truth in code:** [`src/lib/graph-oauth-scopes.ts`](../src/lib/graph-oauth-scopes.ts) (`GRAPH_DEVICE_CODE_LOGIN_SCOPES` for `login`, `GRAPH_REFRESH_SCOPE_CANDIDATES` for token refresh in [`graph-auth`](../src/lib/graph-auth.ts)).

Configure the same permissions on your **Entra ID app registration** (API permissions → Microsoft Graph → Delegated). Then run **`m365-agent-cli login`** so the refresh token includes them. Use **`m365-agent-cli verify-token`** to inspect granted `scp` claims.

**Office 365 Exchange Online:** add **`EWS.AccessAsUser.All`** (delegated) for EWS-backed commands when `M365_EXCHANGE_BACKEND` is `ews` or `auto` (see [`ENTRA_SETUP.md`](./ENTRA_SETUP.md)).

---

## Full scope set (recommended)

| Scope | Purpose in this CLI |
| --- | --- |
| `offline_access` | Refresh tokens |
| `User.Read` | Sign-in profile; `/me` |
| `Calendars.ReadWrite` | Own calendar read/write |
| `Calendars.Read.Shared` | Delegated / shared calendars (`/users/{upn}/calendar/...`) |
| `Calendars.ReadWrite.Shared` | Same, with write |
| `Mail.ReadWrite` | Own mailbox mail APIs |
| `Mail.Read.Shared` | Mail in mailboxes the user can access (delegated/shared) |
| `Mail.ReadWrite.Shared` | Same, with write (where applicable) |
| `MailboxSettings.ReadWrite` | Mailbox settings, OOF, categories, rules-related settings |
| `Place.Read.All` | Places API — `rooms`, room resolution in `create-event` / Places |
| `People.Read` | `GET /me/people` — `find` (people/relevant contacts) |
| `User.Read.All` | `GET /users` directory search — `find` (user query); **often requires admin consent** |
| `Files.ReadWrite.All` | OneDrive / files commands |
| `Sites.ReadWrite.All` | SharePoint / site pages |
| `Tasks.ReadWrite` | Microsoft To Do |
| `Group.ReadWrite.All` | Planner (groups), group-related Graph calls |

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

_Last updated: 2026-04-03 — aligned with `graph-oauth-scopes.ts`._
