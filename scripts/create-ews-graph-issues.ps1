# -----------------------------------------------------------------------------
# WARNING: Do not re-run this script — it creates NEW issues every time.
# It was used once to create GitHub Epic #204 and sub-issues #205–#221.
# Kept as reference for `gh issue create` + GraphQL `addSubIssue`.
# -----------------------------------------------------------------------------
# One-off: create Epic + migration child issues and link as GitHub sub-issues.
# Requires: gh auth, repo: markus-lassfolk/m365-agent-cli
$ErrorActionPreference = "Stop"
$repo = "markus-lassfolk/m365-agent-cli"

function New-MigrationIssue {
  param([string]$Title, [string]$Body, [string[]]$Labels)
  $args = @("issue", "create", "--repo", $repo, "--title", $Title, "--body", $Body)
  foreach ($lb in $Labels) { $args += @("-l", $lb) }
  $url = & gh @args 2>&1
  if ($LASTEXITCODE -ne 0) { throw "gh issue create failed: $url" }
  return ($url | Out-String).Trim()
}

function Add-SubIssue {
  param([string]$ParentNodeId, [string]$ChildIssueUrl)
  $q = 'mutation($p: ID!, $u: String!) { addSubIssue(input: { issueId: $p, subIssueUrl: $u }) { issue { number } subIssue { number } } }'
  $null = gh api graphql -f query=$q -f p=$ParentNodeId -f u=$ChildIssueUrl 2>&1
  if ($LASTEXITCODE -ne 0) { Write-Warning "addSubIssue may have failed for $ChildIssueUrl" }
}

$epicBody = @'
## Goal
Phased migration from **Exchange Web Services (EWS)** to **Microsoft Graph** for Exchange Online, with **Graph as primary** and **EWS as fallback** (`auto` mode) until each slice is verified.

## Tracker
- **[Roadmap, inventory table, phases](https://github.com/markus-lassfolk/m365-agent-cli/blob/main/docs/EWS_TO_GRAPH_MIGRATION_EPIC.md)**
- Issue template: `.github/ISSUE_TEMPLATE/ews-graph-migration.yml`

## References
- [Migrate EWS apps to Microsoft Graph](https://learn.microsoft.com/en-us/graph/migrate-exchange-web-services-overview)
- [EWS to Graph API mapping](https://learn.microsoft.com/en-us/graph/migrate-exchange-web-services-api-mapping)
'@

Write-Host "Creating epic..."
$epicUrl = New-MigrationIssue -Title "Epic: EWS → Microsoft Graph migration (Exchange Online)" -Body $epicBody -Labels @("epic", "migration")
$epicNum = [int]($epicUrl -split "/")[-1]
$epicNodeId = gh api "repos/$repo/issues/$epicNum" --jq .node_id
Write-Host "Epic: $epicUrl (node $epicNodeId)"

$children = @(
  @{
    phase = "Phase 0 — Foundation"
    title = "[Migration] Phase 0 — Foundation (backend router, env, Graph scopes inventory)"
    body  = @"
## Epic
Parent issue: #$epicNum

## Scope
- Agree env vars / ``auto`` fallback (see epic doc)
- Optional: minimal backend router stub (Graph vs EWS)
- Inventory Azure AD app permissions for full Graph mail/calendar/mailboxSettings parity

## Doc
https://github.com/markus-lassfolk/m365-agent-cli/blob/main/docs/EWS_TO_GRAPH_MIGRATION_EPIC.md
"@
  }
  @{
    phase = "Phase 1"
    title = "[Migration] Calendar read — ``calendar`` command (Graph calendarView)"
    body  = @"
## Epic
Parent: #$epicNum

## Inventory
| Area | Commands | Graph direction | Notes |
|------|----------|-----------------|-------|
| Calendar read | ``calendar`` | ``GET calendarView`` / shared calendars | Replace ``getCalendarEvents`` |

## Phase
Phase 1 — Read-only paths
"@
  }
  @{
    phase = "Phase 1"
    title = "[Migration] Free-busy / ``findtime`` — drop EWS ``getScheduleViaOutlook``"
    body  = @"
## Epic
Parent: #$epicNum

## Inventory
| Area | Commands | Graph direction | Notes |
|------|----------|-----------------|-------|
| Free-busy / findtime | ``findtime``, parts of schedule | ``calendar/getSchedule`` | Already partially Graph |

## Phase
Phase 1 — Read-only paths
"@
  }
  @{
    phase = "Phase 1"
    title = "[Migration] ``whoami`` — Graph ``/me`` (drop EWS ResolveNames / getOwaUserInfo)"
    body  = @"
## Epic
Parent: #$epicNum

## Inventory
| Area | Commands | Graph direction | Notes |
|------|----------|-----------------|-------|
| Whoami | ``whoami`` | ``/me`` (+ optional mailboxSettings) | Drop getOwaUserInfo / ResolveNames |

## Phase
Phase 1 — Read-only paths
"@
  }
  @{
    phase = "Phase 2"
    title = "[Migration] Mail — ``mail`` command (list/read/mutations via Graph)"
    body  = @"
## Epic
Parent: #$epicNum

## Inventory
| Area | Commands | Graph direction | Notes |
|------|----------|-----------------|-------|
| Mail CRUD + actions | ``mail`` | Messages, move, patch, send | MIME / large attachment edge cases |

## Phase
Phase 2 — Mail stack
"@
  }
  @{
    phase = "Phase 2"
    title = "[Migration] ``send`` — Graph sendMail"
    body  = @"
## Epic
Parent: #$epicNum

## Phase
Phase 2 — Mail stack
"@
  }
  @{
    phase = "Phase 2"
    title = "[Migration] ``drafts`` — Graph draft messages"
    body  = @"
## Epic
Parent: #$epicNum

## Phase
Phase 2 — Mail stack
"@
  }
  @{
    phase = "Phase 2"
    title = "[Migration] ``folders`` — Graph mailFolders"
    body  = @"
## Epic
Parent: #$epicNum

## Phase
Phase 2 — Mail stack
"@
  }
  @{
    phase = "Phase 2"
    title = "[Migration] ``todo --link`` — Graph get message (replace EWS getEmail)"
    body  = @"
## Epic
Parent: #$epicNum

## Inventory
| Area | Commands | Graph direction | Notes |
|------|----------|-----------------|-------|
| Todo link | ``todo --link`` | get message via Graph | Small change |

## Phase
Phase 2 (or 1) — quick win
"@
  }
  @{
    phase = "Phase 3"
    title = "[Migration] Calendar write — ``create-event``, ``update-event``, ``delete-event``"
    body  = @"
## Epic
Parent: #$epicNum

## Inventory
| Area | Commands | Graph direction | Notes |
|------|----------|-----------------|-------|
| Calendar write | create/update/delete-event | Events API + online meetings | TZ, recurrence, Teams |

## Phase
Phase 3 — Calendar writes + meeting actions
"@
  }
  @{
    phase = "Phase 3"
    title = "[Migration] ``respond`` — accept/decline/tentative via Graph"
    body  = @"
## Epic
Parent: #$epicNum

## Inventory
| Area | Commands | Graph direction | Notes |
|------|----------|-----------------|-------|
| Meeting response | ``respond`` | Graph | Shared mailbox = ``/users/{id}/`` |

## Phase
Phase 3 — Calendar writes + meeting actions
"@
  }
  @{
    phase = "Phase 3"
    title = "[Migration] ``forward-event`` and ``counter`` — Graph equivalents"
    body  = @"
## Epic
Parent: #$epicNum

## Inventory
| Area | Commands | Graph direction | Notes |
|------|----------|-----------------|-------|
| Forward / counter | ``forward-event``, ``counter`` | Verify Graph APIs | |

## Phase
Phase 3 — Calendar writes + meeting actions
"@
  }
  @{
    phase = "Phase 4"
    title = "[Migration] ``auto-reply`` vs ``oof`` — consolidate on Graph mailbox settings"
    body  = @"
## Epic
Parent: #$epicNum

## Inventory
| Area | Commands | Graph direction | Notes |
|------|----------|-----------------|-------|
| Auto-reply (EWS) | ``auto-reply`` | Deprecate; use ``oof`` / mailboxSettings | Align UX |

## Phase
Phase 4 — Rules / OOF consolidation
"@
  }
  @{
    phase = "Phase 5"
    title = "[Migration] ``delegates`` — Graph calendar share/delegate (replace EWS delegate-client)"
    body  = @"
## Epic
Parent: #$epicNum

## Inventory
| Area | Commands | Graph direction | Notes |
|------|----------|-----------------|-------|
| Delegates | ``delegates``, ``delegate-client.ts`` | Calendar permission / share APIs | **No 1:1 EWS matrix** — product redesign |

## Phase
Phase 5 — Delegates (redesign)

## Docs
https://learn.microsoft.com/en-us/graph/outlook-share-or-delegate-calendar
"@
  }
  @{
    phase = "Phase 0 / 6"
    title = "[Migration] Auth — single token cache + Graph scopes (retire EWS_REFRESH_TOKEN path)"
    body  = @"
## Epic
Parent: #$epicNum

## Inventory
| Area | Commands | Graph direction | Notes |
|------|----------|-----------------|-------|
| Auth | ``auth.ts``, env ``EWS_*`` | Single token + Graph scopes | See ``docs/GOALS.md`` |

## Phase
Phase 0 start → Phase 6 finish
"@
  }
  @{
    phase = "Cross-cutting"
    title = "[Migration] Tests and mocks — Graph-shaped fixtures + EWS fallback tests"
    body  = @"
## Epic
Parent: #$epicNum

## Inventory
| Area | Commands | Graph direction | Notes |
|------|----------|-----------------|-------|
| Tests / mocks | ``src/test/mocks``, integration | Graph-shaped mocks | |

## Phase
Ongoing; EWS mocks removed in Phase 6
"@
  }
  @{
    phase = "Phase 6"
    title = "[Migration] Docs — README, ENTRA_SETUP, SKILL (remove EWS after cutover)"
    body  = @"
## Epic
Parent: #$epicNum

## Inventory
| Area | Commands | Graph direction | Notes |
|------|----------|-----------------|-------|
| Docs | README, ENTRA_SETUP, SKILL | Graph-only setup | Last after code cutover |

## Phase
Phase 6 — EWS removal
"@
  }
)

$childUrls = @()
foreach ($c in $children) {
  Write-Host "Creating: $($c.title)"
  $body = $c.body + "`n`n---`n**Phase:** $($c.phase)"
  $url = New-MigrationIssue -Title $c.title -Body $body -Labels @("migration", "ews", "graph")
  $childUrls += $url
  Write-Host "  -> $url"
  Add-SubIssue -ParentNodeId $epicNodeId -ChildIssueUrl $url
  Start-Sleep -Milliseconds 400
}

Write-Host "`nDone."
Write-Host "Epic: $epicUrl"
Write-Host "Sub-issues: $($childUrls.Count)"
