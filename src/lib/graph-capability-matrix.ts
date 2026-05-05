/**
 * Maps Microsoft Graph delegated (scp) or application (roles) permission names on an access token
 * to m365-agent-cli feature areas with read vs write coverage.
 *
 * @see docs/GRAPH_SCOPES.md
 */

export type TokenPermissionKind = 'delegated' | 'application' | 'mixed' | 'unknown';

export interface GraphTokenPayloadForCapabilities {
  scp?: string;
  roles?: string[];
}

export interface CapabilityMatrixRow {
  id: string;
  /** Short group label */
  area: string;
  /** One-line CLI context */
  detail: string;
  /** Scopes that satisfy read access (any one). Write scopes count toward read only when {@link readColumnDash} is false. */
  readScopes: string[];
  /** Scopes that satisfy write/mutate access (any one). */
  writeScopes: string[];
  /** When true, row is informational (e.g. EWS not represented in Graph token). */
  notApplicable?: boolean;
  /** Read column shows — (not a meaningful read/write pair for this area). */
  readColumnDash?: boolean;
  /** Write column shows — */
  writeColumnDash?: boolean;
}

/** Build permission set from JWT payload (`scp` space-separated and/or `roles` array). */
export function permissionSetFromGraphPayload(payload: GraphTokenPayloadForCapabilities): Set<string> {
  const out = new Set<string>();
  if (typeof payload.scp === 'string' && payload.scp.trim()) {
    for (const s of payload.scp.split(/\s+/)) {
      if (s) out.add(s);
    }
  }
  if (Array.isArray(payload.roles)) {
    for (const r of payload.roles) {
      if (typeof r === 'string' && r.trim()) out.add(r.trim());
    }
  }
  return out;
}

export function graphTokenPermissionKind(payload: GraphTokenPayloadForCapabilities): TokenPermissionKind {
  const hasScp = typeof payload.scp === 'string' && payload.scp.trim().length > 0;
  const hasRoles = Array.isArray(payload.roles) && payload.roles.length > 0;
  if (hasScp && hasRoles) return 'mixed';
  if (hasScp) return 'delegated';
  if (hasRoles) return 'application';
  return 'unknown';
}

function hasAnyPermission(perms: Set<string>, candidates: string[]): boolean {
  return candidates.some((c) => perms.has(c));
}

export interface EvaluatedCapability extends CapabilityMatrixRow {
  readOk: boolean;
  writeOk: boolean;
}

function evaluateRow(perms: Set<string>, row: CapabilityMatrixRow): EvaluatedCapability {
  if (row.notApplicable) {
    return { ...row, readOk: false, writeOk: false };
  }
  const readFromRead = hasAnyPermission(perms, row.readScopes);
  const readFromWrite = row.readColumnDash ? false : hasAnyPermission(perms, row.writeScopes);
  const readOk = readFromRead || readFromWrite;
  const writeOk = row.writeColumnDash ? false : hasAnyPermission(perms, row.writeScopes);
  return { ...row, readOk, writeOk };
}

/**
 * Ordered rows for CLI / JSON output. Read is satisfied by any read scope or any listed write scope
 * (e.g. `Calendars.ReadWrite` unlocks calendar read).
 */
export const GRAPH_CAPABILITY_MATRIX: readonly CapabilityMatrixRow[] = [
  {
    id: 'profile',
    area: 'Profile / sign-in',
    detail: '`whoami`, basic `/me`',
    readScopes: ['User.Read', 'User.ReadWrite'],
    writeScopes: ['User.ReadWrite']
  },
  {
    id: 'directory',
    area: 'Directory users',
    detail: '`find` user search — often admin consent',
    readScopes: ['User.Read.All', 'Directory.Read.All', 'Directory.ReadWrite.All'],
    writeScopes: ['Directory.ReadWrite.All']
  },
  {
    id: 'org.relations',
    area: 'Manager & direct reports',
    detail: '`org manager`, `org direct-reports` — self with `User.Read`; other users need `User.Read.All` (typical)',
    readScopes: ['User.Read', 'User.Read.All', 'Directory.Read.All', 'Directory.ReadWrite.All'],
    writeScopes: []
  },
  {
    id: 'calendar.own',
    area: 'Calendar (your mailbox)',
    detail: '`calendar`, `create-event`, `respond`, …',
    readScopes: ['Calendars.Read', 'Calendars.ReadWrite'],
    writeScopes: ['Calendars.ReadWrite']
  },
  {
    id: 'calendar.shared',
    area: 'Calendar (shared / delegated)',
    detail: '`calendar --mailbox`, delegated calendars',
    readScopes: ['Calendars.Read.Shared', 'Calendars.ReadWrite.Shared'],
    writeScopes: ['Calendars.ReadWrite.Shared']
  },
  {
    id: 'mail.own',
    area: 'Mail (your mailbox)',
    detail: '`mail`, `folders`, `drafts` (Graph path)',
    readScopes: ['Mail.Read', 'Mail.ReadWrite'],
    writeScopes: ['Mail.ReadWrite']
  },
  {
    id: 'mail.shared',
    area: 'Mail (shared / delegated)',
    detail: '`mail --mailbox`, shared folders',
    readScopes: ['Mail.Read.Shared', 'Mail.ReadWrite.Shared'],
    writeScopes: ['Mail.ReadWrite.Shared']
  },
  {
    id: 'mail.send',
    area: 'Send mail (Graph)',
    detail: '`send` — `Mail.Send` alone can send; read mail needs `Mail.ReadWrite`',
    readScopes: [],
    writeScopes: ['Mail.Send', 'Mail.ReadWrite'],
    readColumnDash: true
  },
  {
    id: 'mailbox.settings',
    area: 'Mailbox settings',
    detail: '`oof`, categories, mailbox settings',
    readScopes: ['MailboxSettings.Read', 'MailboxSettings.ReadWrite'],
    writeScopes: ['MailboxSettings.ReadWrite']
  },
  {
    id: 'rooms',
    area: 'Rooms & places',
    detail: '`rooms`, Places in `create-event`',
    readScopes: ['Place.Read.All'],
    writeScopes: []
  },
  {
    id: 'people',
    area: 'People / relevance',
    detail: '`find` people, `/me/people`',
    readScopes: ['People.Read', 'People.Read.All'],
    writeScopes: []
  },
  {
    id: 'files.onedrive',
    area: 'OneDrive / files',
    detail:
      '`files` (delta, shared-with-me, thumbnails, copy/move, drive flags, invite, permissions), `excel` workbooks (tables/pivots/ranges/sessions/application) + beta comments, `word`/`powerpoint` preview/meta/download/thumbnails',
    readScopes: ['Files.Read', 'Files.Read.All', 'Files.ReadWrite', 'Files.ReadWrite.All'],
    writeScopes: ['Files.ReadWrite', 'Files.ReadWrite.All']
  },
  {
    id: 'sharepoint.sites',
    area: 'SharePoint sites',
    detail: '`sharepoint` lists/items/create/update/get/delete, `items-delta`, `resolve-site`; `site-pages`',
    readScopes: ['Sites.Read.All', 'Sites.ReadWrite.All', 'Sites.Manage.All'],
    writeScopes: ['Sites.ReadWrite.All', 'Sites.Manage.All']
  },
  {
    id: 'sharepoint.followed',
    area: 'SharePoint followed sites',
    detail: '`sharepoint followed-sites`, `follow`, `unfollow`',
    readScopes: ['Sites.Read.All', 'Sites.ReadWrite.All'],
    writeScopes: ['Sites.ReadWrite.All']
  },
  {
    id: 'insights',
    area: 'Discovery / Insights',
    detail: '`insights trending|used|shared`, `files recent`, `files activities`, `files preview`',
    readScopes: [
      'Sites.Read.All',
      'Sites.ReadWrite.All',
      'Files.Read',
      'Files.Read.All',
      'Files.ReadWrite',
      'Files.ReadWrite.All'
    ],
    writeScopes: []
  },
  {
    id: 'approvals',
    area: 'Approvals',
    detail:
      '`approvals list|get|steps|respond` — Teams Approvals + Power Automate (beta `/me/approvals`); `ApprovalSolution.ReadWrite` (canonical) or narrower `ApprovalSolutionResponse.ReadWrite`',
    readScopes: [
      'ApprovalSolution.Read.All',
      'ApprovalSolution.ReadWrite',
      'ApprovalSolution.ReadWrite.All',
      'ApprovalSolutionResponse.ReadWrite'
    ],
    writeScopes: ['ApprovalSolution.ReadWrite', 'ApprovalSolutionResponse.ReadWrite']
  },
  {
    id: 'todo',
    area: 'Microsoft To Do',
    detail: '`todo`',
    readScopes: ['Tasks.Read', 'Tasks.ReadWrite'],
    writeScopes: ['Tasks.ReadWrite']
  },
  {
    id: 'planner.groups',
    area: 'Planner & group-backed Teams',
    detail: '`planner`, `teams` members/channels/apps/tabs — broad group scope',
    readScopes: ['Group.Read.All', 'Group.ReadWrite.All'],
    writeScopes: ['Group.ReadWrite.All']
  },
  {
    id: 'groups.outlook',
    area: 'Outlook Groups (Microsoft 365 groups)',
    detail: '`groups list`, `conversations`, `thread`, `posts`, `post-reply`',
    readScopes: ['Group.Read.All', 'Group.ReadWrite.All'],
    writeScopes: ['Group.ReadWrite.All']
  },
  {
    id: 'contacts.own',
    area: 'Contacts (your mailbox)',
    detail: '`contacts`',
    readScopes: ['Contacts.Read', 'Contacts.ReadWrite'],
    writeScopes: ['Contacts.ReadWrite']
  },
  {
    id: 'contacts.shared',
    area: 'Contacts (shared mailbox)',
    detail: '`contacts --user`',
    readScopes: ['Contacts.Read.Shared', 'Contacts.ReadWrite.Shared'],
    writeScopes: ['Contacts.ReadWrite.Shared']
  },
  {
    id: 'meetings.online',
    area: 'Online meetings',
    detail: '`meeting`, Teams links in `create-event`',
    readScopes: ['OnlineMeetings.Read', 'OnlineMeetings.ReadWrite'],
    writeScopes: ['OnlineMeetings.ReadWrite']
  },
  {
    id: 'meetings.recordings',
    area: 'Meeting recordings',
    detail:
      '`meeting recordings`, `recording-download`, `recordings-all` (+ `--delta`) — tenant Stream/Teams policy applies',
    readScopes: ['OnlineMeetingRecording.Read.All'],
    writeScopes: []
  },
  {
    id: 'meetings.transcripts',
    area: 'Meeting transcripts',
    detail: '`meeting transcripts`, `transcript-download`, `transcripts-all` (+ `--delta`)',
    readScopes: ['OnlineMeetingTranscript.Read.All'],
    writeScopes: []
  },
  {
    id: 'onenote',
    area: 'OneNote',
    detail: '`onenote`',
    readScopes: ['Notes.Read', 'Notes.ReadWrite', 'Notes.ReadWrite.All'],
    writeScopes: ['Notes.ReadWrite', 'Notes.ReadWrite.All']
  },
  {
    id: 'teams.core',
    area: 'Teams (teams & channels)',
    detail: '`teams` list (incl. `list --user`), channels, `channel-files-folder`, metadata',
    readScopes: ['Team.ReadBasic.All', 'Channel.ReadBasic.All'],
    writeScopes: []
  },
  {
    id: 'teams.channel.messages',
    area: 'Teams channel messages',
    detail: '`teams messages`, read channel posts — often admin consent',
    readScopes: ['ChannelMessage.Read.All'],
    writeScopes: []
  },
  {
    id: 'teams.channel.send',
    area: 'Teams channel send',
    detail:
      '`teams channel-message-send`, replies, `channel-message-patch`, `channel-message-delete` (soft/hard), `tab-create` / `tab-update` / `tab-delete`',
    readScopes: [],
    writeScopes: ['ChannelMessage.Send'],
    readColumnDash: true
  },
  {
    id: 'teams.chats',
    area: 'Teams chats (1:1 / group)',
    detail: '`teams chats`, messages, `chat-create`, `chat-member-add`, `chat-message-patch`, `chat-message-delete`',
    readScopes: ['Chat.Read', 'Chat.ReadWrite'],
    writeScopes: ['Chat.ReadWrite']
  },
  {
    id: 'teams.activity',
    area: 'Teams activity feed',
    detail:
      '`teams activity-notify` — POST /me/teamwork/sendActivityNotification or /chats/{id}/sendActivityNotification',
    readScopes: [],
    writeScopes: ['TeamsActivity.Send'],
    readColumnDash: true
  },
  {
    id: 'teams.members.write',
    area: 'Teams membership (provision)',
    detail: '`teams team-member-add`, `teams channel-member-add`',
    readScopes: [],
    writeScopes: [
      'TeamMember.ReadWrite.All',
      'TeamMember.ReadWriteNonGuestRole.All',
      'ChannelMember.ReadWrite.All',
      'Group.ReadWrite.All'
    ],
    readColumnDash: true
  },
  {
    id: 'presence.read',
    area: 'Presence (read)',
    detail: '`presence me`, `presence user`, bulk',
    readScopes: ['Presence.Read.All'],
    writeScopes: []
  },
  {
    id: 'presence.write',
    area: 'Presence (set/clear)',
    detail: '`presence set-*`, `presence clear-*`',
    readScopes: [],
    writeScopes: ['Presence.ReadWrite'],
    readColumnDash: true
  },
  {
    id: 'bookings',
    area: 'Bookings',
    detail: '`bookings`',
    readScopes: ['Bookings.Read.All', 'Bookings.ReadWrite.All'],
    writeScopes: ['Bookings.ReadWrite.All']
  },
  {
    id: 'graph.search',
    area: 'Graph Search',
    detail: '`graph-search` — needs mail/files/site scopes for each entity type',
    readScopes: [
      'Mail.Read',
      'Mail.ReadWrite',
      'Files.Read.All',
      'Files.ReadWrite.All',
      'Sites.Read.All',
      'Sites.ReadWrite.All'
    ],
    writeScopes: []
  },
  {
    id: 'copilot.retrieval',
    area: 'Copilot Retrieval API',
    detail: '`copilot retrieval` — SharePoint/OneDrive need Files+Sites read; connectors need `ExternalItem.Read.All`',
    readScopes: [
      'Files.Read.All',
      'Files.ReadWrite.All',
      'Sites.Read.All',
      'Sites.ReadWrite.All',
      'ExternalItem.Read.All'
    ],
    writeScopes: []
  },
  {
    id: 'copilot.search',
    area: 'Copilot Search API (preview)',
    detail: '`copilot search` — OneDrive; `Files.Read.All`+`Sites.Read.All` (or ReadWrite equivalents)',
    readScopes: ['Files.Read.All', 'Files.ReadWrite.All', 'Sites.Read.All', 'Sites.ReadWrite.All'],
    writeScopes: []
  },
  {
    id: 'copilot.chat',
    area: 'Copilot Chat API (preview)',
    detail:
      '`copilot conversation-create`, `chat`, `chat-stream` — Microsoft requires the full permission bundle documented for Chat API',
    readScopes: [
      'Sites.Read.All',
      'Sites.ReadWrite.All',
      'Mail.Read',
      'Mail.ReadWrite',
      'Mail.ReadWrite.Shared',
      'People.Read.All',
      'OnlineMeetingTranscript.Read.All',
      'Chat.Read',
      'Chat.ReadWrite',
      'ChannelMessage.Read.All',
      'ExternalItem.Read.All'
    ],
    writeScopes: []
  },
  {
    id: 'copilot.interactions',
    area: 'Copilot interaction export',
    detail: '`copilot interactions-export` — application `AiEnterpriseInteraction.Read.All` (delegated not supported)',
    readScopes: ['AiEnterpriseInteraction.Read.All'],
    writeScopes: []
  },
  {
    id: 'copilot.interactions.notify',
    area: 'Copilot interaction change notifications',
    detail:
      '`subscribe copilot-interactions` — delegated `AiEnterpriseInteraction.Read` (per-user); tenant subscription is app-only',
    readScopes: ['AiEnterpriseInteraction.Read', 'AiEnterpriseInteraction.Read.All'],
    writeScopes: []
  },
  {
    id: 'copilot.meetingInsights',
    area: 'Copilot meeting insights',
    detail: '`copilot meeting-insights-list`, `meeting-insight-get`',
    readScopes: ['OnlineMeetingAiInsight.Read.All'],
    writeScopes: []
  },
  {
    id: 'copilot.reports',
    area: 'Copilot usage reports',
    detail: '`copilot reports …` — `Reports.Read.All` + admin reader role per Microsoft',
    readScopes: ['Reports.Read.All'],
    writeScopes: []
  },
  {
    id: 'copilot.packages',
    area: 'Copilot package catalog (admin)',
    detail:
      '`copilot packages list|get` — `CopilotPackages.Read.All`; `update|block|unblock|reassign` — `CopilotPackages.ReadWrite.All`',
    readScopes: ['CopilotPackages.Read.All', 'CopilotPackages.ReadWrite.All'],
    writeScopes: ['CopilotPackages.ReadWrite.All']
  },
  {
    id: 'graph.invoke',
    area: 'Graph invoke / batch',
    detail: '`graph invoke`, `graph batch` — depends on path you call',
    readScopes: [],
    writeScopes: [],
    readColumnDash: true,
    writeColumnDash: true
  },
  {
    id: 'ews',
    area: 'Exchange Web Services (EWS)',
    detail: 'Not in Graph `scp` — add `EWS.AccessAsUser.All` (Exchange Online) on the same Entra app',
    readScopes: [],
    writeScopes: [],
    notApplicable: true
  }
];

export function evaluateGraphCapabilities(perms: Set<string>): EvaluatedCapability[] {
  return GRAPH_CAPABILITY_MATRIX.map((row) => evaluateRow(perms, row));
}

function checkbox(ok: boolean): string {
  return ok ? '[x]' : '[ ]';
}

/** Human-readable table: Feature | Read | Write (Write shows — when no write scopes apply). */
export function formatCapabilityTextTable(evaluated: EvaluatedCapability[], opts?: { verbose?: boolean }): string {
  const colW = 38;
  const lines: string[] = [];
  lines.push(`${'Feature'.padEnd(colW)} Read   Write`);
  lines.push('─'.repeat(colW + 14));
  for (const r of evaluated) {
    let readCell: string;
    if (r.notApplicable || r.readColumnDash) {
      readCell = ' —  ';
    } else {
      readCell = `${checkbox(r.readOk)} `.padEnd(5);
    }
    let writeCell: string;
    if (r.notApplicable || r.writeColumnDash) {
      writeCell = ' —';
    } else if (r.writeScopes.length === 0) {
      writeCell = ' —';
    } else {
      writeCell = checkbox(r.writeOk);
    }
    lines.push(`${r.area.padEnd(colW)} ${readCell} ${writeCell}`);
    if (opts?.verbose) {
      lines.push(`  ${r.detail}`);
    }
  }
  return lines.join('\n');
}
