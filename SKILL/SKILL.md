---
name: clippy
description: Microsoft 365 / Outlook CLI using EWS SOAP API + OAuth2. Manage calendar (view, create, update, delete events, find meeting times, respond to invitations), send/read/search email, shared mailbox support, and OneDrive file operations via Microsoft Graph.
metadata: {"clawdbot":{"requires":{"bins":["clippy"]}}}
---

# Clippy - Microsoft 365 CLI (EWS Edition)

Source: https://github.com/markus-lassfolk/clippy

Uses EWS SOAP API with OAuth2 refresh token auth. Runs on Bun.

## Install

```bash
git clone https://github.com/markus-lassfolk/clippy.git
cd clippy && bun install
bun run src/cli.ts --help
```

## Auth Setup

Set these env vars (e.g. in `.env`):

```bash
EWS_CLIENT_ID=<Azure AD app client ID>
EWS_REFRESH_TOKEN=<OAuth2 refresh token>
EWS_USERNAME=<your email>
EWS_ENDPOINT=https://outlook.office365.com/EWS/Exchange.asmx
EWS_TENANT_ID=common  # or your tenant ID
```

For shared mailbox access:
```bash
EWS_TARGET_MAILBOX=<shared@mailbox.com>
```

Check auth: `bun run src/cli.ts whoami`

## Commands

### Calendar

```bash
clippy calendar                        # today's events
clippy calendar --day tomorrow
clippy calendar --week
clippy calendar --details

# Write ops
clippy create-event "Meeting" 14:00 15:00 --day tomorrow --description "Notes"
clippy update-event <id> --title "New Title"
clippy delete-event <id>
clippy respond accept --id <eventId>
clippy findtime --attendees "a@co.com,b@co.com" --duration 60
```

### Shared Mailbox Calendar (--mailbox flag)

```bash
clippy calendar --mailbox shared@company.com
clippy create-event "Team Meeting" 10:00 11:00 --mailbox shared@company.com
```

### Email

```bash
clippy mail                            # inbox
clippy mail sent
clippy mail -r <number>               # read email
clippy mail --search "invoice"

# Write ops
clippy send --to "recipient@company.com" --subject "Hello" --body "Body"
clippy mail --reply <number> --message "Thanks!"
clippy mail --forward <number> --to-addr "colleague@company.com"
```

### Shared Mailbox Email (--mailbox flag)

```bash
clippy mail --mailbox shared@company.com
clippy send --to "recipient@company.com" --subject "From shared" --body "..." --mailbox shared@company.com
```

### Other

```bash
clippy folders                         # list mail folders
clippy find "john"                    # people search
clippy findtime                        # find meeting slots
clippy whoami                          # check auth
```

## Architecture

- **EWS** (`src/ews-client.ts`): SOAP calls to Exchange Online
- **Auth** (`src/auth.ts`): OAuth2 refresh token, token cache at `~/.config/clippy/token-cache.json`
- **Commands** (`src/commands/`): `mail.ts`, `calendar.ts`, `send.ts`, `create-event.ts`, etc.
