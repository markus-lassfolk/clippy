#!/usr/bin/env npx tsx
/**
 * Regenerate docs/GRAPH_PERMISSION_FEATURE_MATRIX.md from GRAPH_CAPABILITY_MATRIX.
 * Run from repo root: npm run docs:graph-permission-matrix
 * CI / drift check: npm run docs:graph-permission-matrix:check
 */
import { readFileSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { type CapabilityMatrixRow, GRAPH_CAPABILITY_MATRIX } from '../src/lib/graph-capability-matrix.ts';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const outPath = join(root, 'docs', 'GRAPH_PERMISSION_FEATURE_MATRIX.md');

/** Optional §2 “Notes” column; keyed by exact permission string from the matrix. */
const PERMISSION_NOTES: Record<string, string> = {
  'AiEnterpriseInteraction.Read': 'Delegated per-user subscriptions',
  'AiEnterpriseInteraction.Read.All': 'Export often **application**; `.All` also listed for notifications in matrix',
  'ApprovalSolution.Read.All': 'Read/list steps',
  'ApprovalSolution.ReadWrite': 'Create/respond; canonical name in [`GRAPH_SCOPES.md`](./GRAPH_SCOPES.md)',
  'ApprovalSolution.ReadWrite.All': 'Read side in matrix',
  'ApprovalSolutionResponse.ReadWrite': 'Narrower respond-only variant',
  'Bookings.Read.All': 'Read',
  'Bookings.ReadWrite.All': 'Read + write',
  'Calendars.Read': 'Read',
  'Calendars.ReadWrite': 'Read + write',
  'Calendars.Read.Shared': 'Read',
  'Calendars.ReadWrite.Shared': 'Read + write',
  'Channel.ReadBasic.All': 'Channels metadata',
  'ChannelMember.ReadWrite.All': 'Add channel members',
  'ChannelMessage.Read.All': 'Often admin consent',
  'ChannelMessage.Send': 'Post/reply/patch/delete channel messages & tab CRUD per matrix',
  'Chat.Read': 'Read',
  'Chat.ReadWrite': 'Read + write chats',
  'Contacts.Read': 'Read',
  'Contacts.ReadWrite': 'Read + write',
  'Contacts.Read.Shared': 'Read',
  'Contacts.ReadWrite.Shared': 'Read + write',
  'CopilotPackages.Read.All': 'Read',
  'CopilotPackages.ReadWrite.All': 'Mutations',
  'Directory.Read.All': 'With `User.Read.All`-style directory reads',
  'Directory.ReadWrite.All': 'Write',
  'ExternalItem.Read.All': 'Connectors / external items',
  'Files.Read': 'Narrower file read',
  'Files.Read.All': 'All drives user can reach',
  'Files.ReadWrite': 'File write',
  'Files.ReadWrite.All': 'Broad file read/write',
  'Group.Read.All': 'Read',
  'Group.ReadWrite.All': 'Includes member add paths per matrix',
  'Mail.Read': 'Read',
  'Mail.ReadWrite': 'Read + write mailbox',
  'Mail.Read.Shared': 'Read',
  'Mail.ReadWrite.Shared': 'Shared mailbox write',
  'Mail.Send': 'Sending without full mailbox read',
  'MailboxSettings.Read': 'Read',
  'MailboxSettings.ReadWrite': 'Read + write',
  'Notes.Read': 'Read',
  'Notes.ReadWrite': 'Read + write',
  'Notes.ReadWrite.All': 'All notebooks',
  'OnlineMeetingRecording.Read.All': 'Tenant Stream/Teams policy may still block',
  'OnlineMeetings.Read': 'Read',
  'OnlineMeetings.ReadWrite': 'Create/update/delete',
  'People.Read': '`/me/people`, etc.',
  'People.Read.All': 'Directory-style people reads',
  'Place.Read.All': 'Often admin consent',
  'Reports.Read.All': 'Plus Microsoft’s reports reader role',
  'Sites.Manage.All': 'Elevated site manage',
  'Team.ReadBasic.All': 'Joined teams',
  'TeamMember.ReadWrite.All': 'Add team members',
  'TeamMember.ReadWriteNonGuestRole.All': 'Non-guest variant',
  'TeamsActivity.Send': 'Activity notifications',
  'User.Read.All': 'Often admin consent'
};

function backtickScopes(scopes: readonly string[]): string {
  return scopes.map((s) => `\`${s}\``).join(', ');
}

function featureReadCell(row: CapabilityMatrixRow): string {
  if (row.docFeatureMatrixReadCell !== undefined) return row.docFeatureMatrixReadCell;
  if (row.notApplicable) return '*Not in Graph `scp`*';
  if (row.readColumnDash) return '—';
  return backtickScopes(row.readScopes) + (row.docFeatureMatrixReadSuffix ?? '');
}

function featureWriteCell(row: CapabilityMatrixRow): string {
  if (row.docFeatureMatrixWriteCell !== undefined) return row.docFeatureMatrixWriteCell;
  if (row.notApplicable) return '*Not in Graph `scp`* — add `EWS.AccessAsUser.All`';
  if (row.writeColumnDash || row.writeScopes.length === 0) return '—';
  return backtickScopes(row.writeScopes);
}

function escapeTableCell(s: string): string {
  return s.replace(/\|/g, '\\|').replace(/\n/g, ' ');
}

function buildPermissionToAreas(): Map<string, Set<string>> {
  const map = new Map<string, Set<string>>();
  for (const row of GRAPH_CAPABILITY_MATRIX) {
    if (row.notApplicable) continue;
    const scopes = [...row.readScopes, ...row.writeScopes];
    for (const p of scopes) {
      let set = map.get(p);
      if (!set) {
        set = new Set();
        map.set(p, set);
      }
      set.add(row.area);
    }
  }
  return map;
}

function generatedAt(): string {
  return new Date().toISOString().slice(0, 10);
}

/** Normalize footer so `--check` compares matrix content only (not calendar day). */
function normalizeMatrixDocForCompare(markdown: string): string {
  return markdown.replace(/Generated: \d{4}-\d{2}-\d{2}/g, 'Generated: <date>');
}

function emit(): string {
  const permToAreas = buildPermissionToAreas();
  const sortedPerms = [...permToAreas.keys()].sort((a, b) => a.localeCompare(b));

  const table1Rows = GRAPH_CAPABILITY_MATRIX.map((row) => {
    const area = escapeTableCell(row.area);
    const detail = escapeTableCell(row.detail);
    const read = escapeTableCell(featureReadCell(row));
    const write = escapeTableCell(featureWriteCell(row));
    return `| ${area} | ${detail} | ${read} | ${write} |`;
  });

  const table2Rows = sortedPerms.map((perm) => {
    const areas = [...permToAreas.get(perm)!].sort((a, b) => a.localeCompare(b));
    const enables = areas.join('; ');
    const note = PERMISSION_NOTES[perm] ?? '';
    return `| \`${perm}\` | ${escapeTableCell(enables)} | ${escapeTableCell(note)} |`;
  });

  return `# Graph permissions ↔ CLI features (Entra admin matrix)

**Purpose:** Help Entra administrators choose **delegated** Microsoft Graph permissions based on which **m365-agent-cli** capabilities should work, and to see **which features each permission unlocks**.

**Sources of truth**

- **Feature ↔ scope logic (read/write evaluation):** [\`src/lib/graph-capability-matrix.ts\`](../src/lib/graph-capability-matrix.ts) (\`GRAPH_CAPABILITY_MATRIX\`) — also drives **\`m365-agent-cli verify-token --capabilities\`**. This file is **generated** from that matrix via **\`npm run docs:graph-permission-matrix\`**.
- **Scopes requested on \`login\` / refresh:** [\`src/lib/graph-oauth-scopes.ts\`](../src/lib/graph-oauth-scopes.ts).
- **Narrative scope guide:** [\`GRAPH_SCOPES.md\`](./GRAPH_SCOPES.md).
- **Exchange Web Services:** not represented in Graph \`scp\`; add **\`EWS.AccessAsUser.All\`** (Exchange Online) when using EWS-backed mail/calendar — see [\`ENTRA_SETUP.md\`](./ENTRA_SETUP.md).

**How to verify after consent:** \`m365-agent-cli verify-token\` (inspect \`scp\`) and **\`verify-token --capabilities\`** (checklist). Add **\`--json\`** for automation.

**Legend**

- **Read:** user can use read-only flows for that area if the token includes **any** listed scope (a **Write** scope also satisfies **Read** for that row, unless the row marks read as not applicable).
- **Write:** user can mutate data if the token includes **any** listed scope. **—** means the row is read-only, send-only, or otherwise has no separate “write” column meaning.
- **Least privilege:** prefer narrower permissions where Graph allows; this table lists what the CLI’s capability checker understands, not every possible Graph alternative.

---

## 1) Feature / capability → Graph permissions (pick scopes by product need)

| Feature area | CLI context (summary) | Read — grant one or more | Write — grant one or more |
| --- | --- | --- | --- |
${table1Rows.join('\n')}

---

## 2) Graph permission → features (pick features enabled by each consent)

Alphabetical **delegated** (and noted **application**) permissions referenced by the capability matrix. “Features” names match §1 **Feature area** column.

| Permission | Enables (feature areas) | Notes |
| --- | --- | --- |
${table2Rows.join('\n')}

---

## Related

- [\`GRAPH_PRODUCT_PARITY_MATRIX.md\`](./GRAPH_PRODUCT_PARITY_MATRIX.md) — workloads vs CLI commands (coverage, not permissions).
- [\`GRAPH_INVOKE_BOUNDARIES.md\`](./GRAPH_INVOKE_BOUNDARIES.md) — raw **\`graph invoke\`** surfaces; consent must match each API.
- [\`GRAPH_TROUBLESHOOTING.md\`](./GRAPH_TROUBLESHOOTING.md)

*Auto-generated from \`GRAPH_CAPABILITY_MATRIX\` — run \`npm run docs:graph-permission-matrix\` after editing the matrix. Generated: ${generatedAt()}.*
`;
}

const check = process.argv.includes('--check');
const next = emit();
if (check) {
  let current: string;
  try {
    current = readFileSync(outPath, 'utf8');
  } catch {
    console.error('::error::GRAPH_PERMISSION_FEATURE_MATRIX.md is missing — run npm run docs:graph-permission-matrix');
    process.exit(1);
  }
  if (normalizeMatrixDocForCompare(current) !== normalizeMatrixDocForCompare(next)) {
    console.error(
      '::error::GRAPH_PERMISSION_FEATURE_MATRIX.md is out of sync with GRAPH_CAPABILITY_MATRIX. Run: npm run docs:graph-permission-matrix (then commit the updated doc).'
    );
    console.error(
      `${outPath} is out of date. Run: npm run docs:graph-permission-matrix\n` +
        '(Or apply the diff after editing GRAPH_CAPABILITY_MATRIX / PERMISSION_NOTES in the generator.)'
    );
    process.exit(1);
  }
  console.log(`OK — ${outPath} matches GRAPH_CAPABILITY_MATRIX`);
  process.exit(0);
}

writeFileSync(outPath, next, 'utf8');
console.log(`Wrote ${outPath}`);
