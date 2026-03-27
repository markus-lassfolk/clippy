---
name: clippy
description: Microsoft 365 / Outlook CLI for calendar and email, with Lassfolk-specific read-only target mailbox support via `--mailbox` and `EWS_TARGET_MAILBOX`.
---

# Clippy - Microsoft 365 CLI

Fork: `https://github.com/markus-lassfolk/clippy`

Detta repo är den modifierade Lassfolk-forken av Clippy för Doris/Family Office-flöden.

## Vad som är nytt i denna fork

Denna variant kör EWS/OAuth lokalt och har **read-only-stöd för target/shared mailbox**.

Stöd finns via:
- env: `EWS_TARGET_MAILBOX`
- CLI-flagga: `--mailbox <email>`

Detta gäller i första hand read-only-paths som:
- `mail`
- `calendar`
- `findtime`
- `find`

Målet är att Doris ska kunna vara inloggad som sig själv, men läsa exempelvis `lotta@lassfolk.net` när rättigheter finns i Microsoft 365.

## Lokal körning

```bash
bun run src/cli.ts --help
```

Eller via wrapper om sådan finns lokalt:

```bash
clippy --help
```

## Authentication

Den här EWS-versionen använder OAuth2 refresh token i miljön/cache, inte äldre browser-login-kommandon.

Typiska variabler:

```bash
export EWS_USERNAME="doris@lassfolk.net"
export EWS_CLIENT_ID="..."
export EWS_REFRESH_TOKEN="..."
export EWS_ENDPOINT="https://outlook.office365.com/EWS/Exchange.asmx"
```

Valfritt för default target mailbox:

```bash
export EWS_TARGET_MAILBOX="lotta@lassfolk.net"
```

## Exempel: Doris egen mailbox

```bash
bun run src/cli.ts mail inbox -n 10
bun run src/cli.ts calendar week
```

## Exempel: Lottas mailbox via Doris-auth (read-only)

```bash
bun run src/cli.ts mail inbox --mailbox lotta@lassfolk.net
bun run src/cli.ts mail inbox -n 10 --mailbox lotta@lassfolk.net
bun run src/cli.ts calendar week --mailbox lotta@lassfolk.net
bun run src/cli.ts findtime --mailbox lotta@lassfolk.net
```

Eller med default target mailbox i env:

```bash
export EWS_TARGET_MAILBOX="lotta@lassfolk.net"
bun run src/cli.ts mail inbox
bun run src/cli.ts calendar week
```

## Viktig begränsning

`--mailbox` / `EWS_TARGET_MAILBOX` är i denna fork avsett för **read-only-stöd först**.

Var försiktig med write-paths som:
- send
- move
- mark-read
- update-event
- delete-event
- reply

De är inte huvudsyftet med target-mailbox-funktionen i denna första version.

## Rekommenderad verifiering

```bash
bun install
bun run typecheck
bun run src/cli.ts mail --help
bun run src/cli.ts calendar --help
bun run src/cli.ts mail inbox --mailbox lotta@lassfolk.net
bun run src/cli.ts calendar week --mailbox lotta@lassfolk.net
```

## Why this fork exists

Upstream Clippy är bra, men Doris/Lassfolk behöver ett tydligare arbetsflöde där:
- Doris har eget konto
- Doris kan läsa Lottas mailbox och kalender
- detta sker säkert, kontrollerat och read-only först
- beteendet är stabilt för automation och sekreterarflöden
