import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { Command } from 'commander';
import {
  assertCopilotReportPeriod,
  buildCopilotRetrievalBody,
  buildCopilotSearchBody,
  copilotConversationChat,
  copilotConversationChatOverStream,
  copilotConversationCreate,
  copilotInteractionsExportList,
  copilotMeetingInsightGet,
  copilotMeetingInsightsList,
  copilotPackagesBlock,
  copilotPackagesGet,
  copilotPackagesList,
  copilotPackagesReassign,
  copilotPackagesUnblock,
  copilotPackagesUpdate,
  copilotReportGet,
  copilotRetrieval,
  copilotSearch,
  copilotSearchNextPage,
  COPILOT_RETRIEVAL_DATA_SOURCES
} from '../lib/copilot-graph-client.js';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import { checkReadOnly } from '../lib/utils.js';

export const copilotCommand = new Command('copilot').description(
  'Microsoft 365 Copilot APIs on Microsoft Graph (/copilot/...). Licensing, roles, and preview terms apply; see https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/copilot-apis-overview'
);

type AuthOpts = { token?: string; identity?: string };

async function resolveTokenOrExit(opts: AuthOpts): Promise<string> {
  const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
  if (!auth.success || !auth.token) {
    console.error(`Auth error: ${auth.error}`);
    process.exit(1);
  }
  return auth.token;
}

function printJson(data: unknown): void {
  console.log(JSON.stringify(data, null, 2));
}

function exitGraphError(prefix: string, message: string | undefined): never {
  console.error(`${prefix}${message || 'Unknown error'}`);
  process.exit(1);
}

/** POST /copilot/retrieval */
copilotCommand
  .command('retrieval')
  .description('POST /copilot/retrieval — grounding extracts (SharePoint, OneDrive, connectors)')
  .option('-q, --query <text>', 'Natural language query (max 1500 chars; required with --data-source unless --json-file)')
  .option('-s, --data-source <source>', `With --query: ${COPILOT_RETRIEVAL_DATA_SOURCES.join(' | ')}`)
  .option('--filter-expression <kql>', 'Optional KQL filterExpression')
  .option('--max <n>', 'maximumNumberOfResults (1–25)', (v) => parseInt(String(v), 10))
  .option('-m, --metadata <fields>', 'Comma-separated resourceMetadata names')
  .option('-f, --json-file <path>', 'Full JSON body (overrides query flags)')
  .option('--beta', 'Use Graph beta')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (opts: {
      query?: string;
      dataSource?: string;
      filterExpression?: string;
      max?: number;
      metadata?: string;
      jsonFile?: string;
      beta?: boolean;
      token?: string;
      identity?: string;
    }) => {
      const token = await resolveTokenOrExit(opts);
      let body: Record<string, unknown>;
      if (opts.jsonFile) {
        const raw = await readFile(resolve(process.cwd(), opts.jsonFile.trim()), 'utf8');
        body = JSON.parse(raw) as Record<string, unknown>;
      } else {
        try {
          body = buildCopilotRetrievalBody({
            queryString: opts.query ?? '',
            dataSource: opts.dataSource ?? '',
            filterExpression: opts.filterExpression,
            maximumNumberOfResults: opts.max,
            resourceMetadata: opts.metadata
              ?.split(',')
              .map((s) => s.trim())
              .filter(Boolean)
          });
        } catch (e) {
          console.error(e instanceof Error ? e.message : String(e));
          process.exit(1);
        }
      }
      const r = await copilotRetrieval(token, body, Boolean(opts.beta));
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      printJson(r.data);
    }
  );

/** POST /copilot/search (preview; defaults to beta) */
copilotCommand
  .command('search')
  .description('POST /copilot/search — hybrid search over OneDrive for work or school (preview; beta by default)')
  .option('-q, --query <text>', 'Natural language query (required unless --json-file)')
  .option('--page-size <n>', 'Results per page (1–100)', (v) => parseInt(String(v), 10))
  .option('--one-drive-filter <kql>', 'dataSources.oneDrive.filterExpression (path KQL)')
  .option('-m, --metadata <fields>', 'Comma-separated resourceMetadataNames for OneDrive')
  .option('-f, --json-file <path>', 'Full JSON body')
  .option('--v1', 'Use Graph v1.0 (search is generally beta; v1 may 404 until GA)')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (opts: {
      query?: string;
      pageSize?: number;
      oneDriveFilter?: string;
      metadata?: string;
      jsonFile?: string;
      v1?: boolean;
      token?: string;
      identity?: string;
    }) => {
      const token = await resolveTokenOrExit(opts);
      let body: Record<string, unknown>;
      if (opts.jsonFile) {
        const raw = await readFile(resolve(process.cwd(), opts.jsonFile.trim()), 'utf8');
        body = JSON.parse(raw) as Record<string, unknown>;
      } else {
        try {
          body = buildCopilotSearchBody({
            query: opts.query ?? '',
            pageSize: opts.pageSize,
            oneDriveFilterExpression: opts.oneDriveFilter,
            resourceMetadataNames: opts.metadata
              ?.split(',')
              .map((s) => s.trim())
              .filter(Boolean)
          });
        } catch (e) {
          console.error(e instanceof Error ? e.message : String(e));
          process.exit(1);
        }
      }
      const r = await copilotSearch(token, body, !opts.v1);
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      printJson(r.data);
    }
  );

/** GET @odata.nextLink from search */
copilotCommand
  .command('search-next')
  .description('GET Copilot search next page — pass full @odata.nextLink URL from a prior search response')
  .requiredOption('--next-link <url>', 'Full HTTPS nextLink from copilot search response')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { nextLink: string; token?: string; identity?: string }) => {
    const token = await resolveTokenOrExit(opts);
    const r = await copilotSearchNextPage(token, opts.nextLink);
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    printJson(r.data);
  });

/** POST /copilot/conversations */
copilotCommand
  .command('conversation-create')
  .description('POST /copilot/conversations — create an empty Copilot chat conversation (returns id)')
  .option('--v1', 'Use Graph v1.0 (default is beta, as in Microsoft docs)')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { v1?: boolean; token?: string; identity?: string }, cmd) => {
    checkReadOnly(cmd);
    const token = await resolveTokenOrExit(opts);
    const r = await copilotConversationCreate(token, !opts.v1);
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    printJson(r.data);
  });

/** POST .../chat */
copilotCommand
  .command('chat')
  .description('POST /copilot/conversations/{id}/chat — synchronous Copilot message (requires locationHint)')
  .argument('<conversationId>', 'Conversation id from conversation-create')
  .option('-m, --message <text>', 'User message text (required unless --json-file)')
  .option('-z, --timezone <iana>', 'locationHint.timeZone (e.g. America/New_York; required with --message unless --json-file)')
  .option('-f, --json-file <path>', 'Full JSON body (message, locationHint, contextualResources, …)')
  .option('--v1', 'Use Graph v1.0 (default is beta)')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      conversationId: string,
      opts: {
        message?: string;
        timezone?: string;
        jsonFile?: string;
        v1?: boolean;
        token?: string;
        identity?: string;
      },
      cmd
    ) => {
      checkReadOnly(cmd);
      const token = await resolveTokenOrExit(opts);
      let body: Record<string, unknown>;
      if (opts.jsonFile) {
        const raw = await readFile(resolve(process.cwd(), opts.jsonFile.trim()), 'utf8');
        body = JSON.parse(raw) as Record<string, unknown>;
      } else {
        const text = (opts.message ?? '').trim();
        const tz = (opts.timezone ?? '').trim();
        if (!text || !tz) {
          console.error('Error: --message and --timezone are required unless --json-file is set');
          process.exit(1);
        }
        body = {
          message: { text },
          locationHint: { timeZone: tz }
        };
      }
      const r = await copilotConversationChat(token, conversationId, body, !opts.v1);
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      printJson(r.data);
    }
  );

/** POST .../chatOverStream */
copilotCommand
  .command('chat-stream')
  .description('POST /copilot/conversations/{id}/chatOverStream — streamed SSE response (printed as raw text)')
  .argument('<conversationId>', 'Conversation id')
  .option('-m, --message <text>', 'User message (required unless --json-file)')
  .option('-z, --timezone <iana>', 'locationHint.timeZone (required with --message unless --json-file)')
  .option('-f, --json-file <path>', 'Full JSON body')
  .option('--v1', 'Use Graph v1.0 (default is beta)')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (
      conversationId: string,
      opts: {
        message?: string;
        timezone?: string;
        jsonFile?: string;
        v1?: boolean;
        token?: string;
        identity?: string;
      },
      cmd
    ) => {
      checkReadOnly(cmd);
      const token = await resolveTokenOrExit(opts);
      let body: Record<string, unknown>;
      if (opts.jsonFile) {
        const raw = await readFile(resolve(process.cwd(), opts.jsonFile.trim()), 'utf8');
        body = JSON.parse(raw) as Record<string, unknown>;
      } else {
        const text = (opts.message ?? '').trim();
        const tz = (opts.timezone ?? '').trim();
        if (!text || !tz) {
          console.error('Error: --message and --timezone are required unless --json-file is set');
          process.exit(1);
        }
        body = {
          message: { text },
          locationHint: { timeZone: tz }
        };
      }
      const r = await copilotConversationChatOverStream(token, conversationId, body, !opts.v1);
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      console.log(r.data ?? '');
    }
  );

/** GET interaction export (application permission typical) */
copilotCommand
  .command('interactions-export')
  .description(
    'GET .../interactionHistory/getAllEnterpriseInteractions — export Copilot interactions for a user (app-only AiEnterpriseInteraction.Read.All typical; see Graph docs)'
  )
  .requiredOption('--user <id>', 'User id (GUID or UPN) in the path')
  .option('--odata <query>', 'OData query without leading ? (e.g. $top=100&$filter=... )')
  .option('--beta', 'Use Graph beta')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { user: string; odata?: string; beta?: boolean; token?: string; identity?: string }) => {
    const token = await resolveTokenOrExit(opts);
    const r = await copilotInteractionsExportList(token, opts.user, opts.odata, Boolean(opts.beta));
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    printJson(r.data);
  });

/** GET meeting AI insights list */
copilotCommand
  .command('meeting-insights-list')
  .description('GET /copilot/users/{user}/onlineMeetings/{meetingId}/aiInsights')
  .requiredOption('--user <id>', 'User id')
  .requiredOption('--meeting <id>', 'Online meeting id')
  .option('--odata <query>', 'OData without leading ? (e.g. $select=id)')
  .option('--beta', 'Use Graph beta')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (opts: { user: string; meeting: string; odata?: string; beta?: boolean; token?: string; identity?: string }) => {
      const token = await resolveTokenOrExit(opts);
      const r = await copilotMeetingInsightsList(token, opts.user, opts.meeting, opts.odata, Boolean(opts.beta));
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      printJson(r.data);
    }
  );

/** GET single meeting insight */
copilotCommand
  .command('meeting-insight-get')
  .description('GET /copilot/users/{user}/onlineMeetings/{meetingId}/aiInsights/{insightId}')
  .requiredOption('--user <id>', 'User id')
  .requiredOption('--meeting <id>', 'Online meeting id')
  .requiredOption('--insight <id>', 'AI insight id')
  .option('--odata <query>', 'OData without leading ? (e.g. $select=meetingNotes)')
  .option('--beta', 'Use Graph beta')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (opts: {
      user: string;
      meeting: string;
      insight: string;
      odata?: string;
      beta?: boolean;
      token?: string;
      identity?: string;
    }) => {
      const token = await resolveTokenOrExit(opts);
      const r = await copilotMeetingInsightGet(
        token,
        opts.user,
        opts.meeting,
        opts.insight,
        opts.odata,
        Boolean(opts.beta)
      );
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      printJson(r.data);
    }
  );

const reportsCmd = new Command('reports').description(
  'Copilot usage reports (GET /copilot/reports/...); requires Reports.Read.All and admin report reader roles where applicable'
);

reportsCmd.addCommand(
  new Command('user-count-summary')
    .description('getMicrosoft365CopilotUserCountSummary(period=...)')
    .requiredOption('-p, --period <code>', 'D7 | D30 | D90 | D180 | ALL')
    .option('--beta', 'Use Graph beta')
    .option('--v1', 'Use Graph v1.0')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity')
    .action(async (opts: { period: string; beta?: boolean; v1?: boolean; token?: string; identity?: string }) => {
      const token = await resolveTokenOrExit(opts);
      let period: string;
      try {
        period = assertCopilotReportPeriod(opts.period);
      } catch (e) {
        console.error(e instanceof Error ? e.message : String(e));
        process.exit(1);
      }
      const beta = Boolean(opts.beta) && !opts.v1;
      const r = await copilotReportGet(token, 'getMicrosoft365CopilotUserCountSummary', period, beta);
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      printJson(r.data);
    })
);

reportsCmd.addCommand(
  new Command('user-count-trend')
    .description('getMicrosoft365CopilotUserCountTrend(period=...)')
    .requiredOption('-p, --period <code>', 'D7 | D30 | D90 | D180 | ALL')
    .option('--beta', 'Use Graph beta')
    .option('--v1', 'Use Graph v1.0')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity')
    .action(async (opts: { period: string; beta?: boolean; v1?: boolean; token?: string; identity?: string }) => {
      const token = await resolveTokenOrExit(opts);
      let period: string;
      try {
        period = assertCopilotReportPeriod(opts.period);
      } catch (e) {
        console.error(e instanceof Error ? e.message : String(e));
        process.exit(1);
      }
      const beta = Boolean(opts.beta) && !opts.v1;
      const r = await copilotReportGet(token, 'getMicrosoft365CopilotUserCountTrend', period, beta);
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      printJson(r.data);
    })
);

reportsCmd.addCommand(
  new Command('usage-user-detail')
    .description('getMicrosoft365CopilotUsageUserDetail(period=...)')
    .requiredOption('-p, --period <code>', 'D7 | D30 | D90 | D180 | ALL')
    .option('--beta', 'Use Graph beta')
    .option('--v1', 'Use Graph v1.0')
    .option('--token <token>', 'Graph access token')
    .option('--identity <name>', 'Graph token cache identity')
    .action(async (opts: { period: string; beta?: boolean; v1?: boolean; token?: string; identity?: string }) => {
      const token = await resolveTokenOrExit(opts);
      let period: string;
      try {
        period = assertCopilotReportPeriod(opts.period);
      } catch (e) {
        console.error(e instanceof Error ? e.message : String(e));
        process.exit(1);
      }
      const beta = Boolean(opts.beta) && !opts.v1;
      const r = await copilotReportGet(token, 'getMicrosoft365CopilotUsageUserDetail', period, beta);
      if (!r.ok) exitGraphError('Error: ', r.error?.message);
      printJson(r.data);
    })
);

copilotCommand.addCommand(reportsCmd);

const packagesCmd = new Command('packages').description(
  'Copilot package catalog (beta /copilot/admin/catalog/packages); CopilotPackages.Read*. See Microsoft Agent 365 licensing.'
);

packagesCmd
  .command('list')
  .description('GET /copilot/admin/catalog/packages')
  .option('--odata <query>', 'OData without leading ?')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { odata?: string; token?: string; identity?: string }) => {
    const token = await resolveTokenOrExit(opts);
    const r = await copilotPackagesList(token, opts.odata);
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    printJson(r.data);
  });

packagesCmd
  .command('get')
  .description('GET /copilot/admin/catalog/packages/{id}')
  .argument('<packageId>', 'Package id (e.g. P_...)')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (packageId: string, opts: AuthOpts) => {
    const token = await resolveTokenOrExit(opts);
    const r = await copilotPackagesGet(token, packageId);
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    printJson(r.data);
  });

packagesCmd
  .command('update')
  .description('PATCH /copilot/admin/catalog/packages/{id}')
  .argument('<packageId>', 'Package id')
  .requiredOption('-f, --json-file <path>', 'JSON body (allowedUsersAndGroups, acquireUsersAndGroups, …)')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (packageId: string, opts: AuthOpts & { jsonFile: string }, cmd) => {
    checkReadOnly(cmd);
    const token = await resolveTokenOrExit(opts);
    const raw = await readFile(resolve(process.cwd(), opts.jsonFile.trim()), 'utf8');
    const body = JSON.parse(raw) as Record<string, unknown>;
    const r = await copilotPackagesUpdate(token, packageId, body);
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    if (r.data !== undefined) printJson(r.data);
    else console.log('OK (204 No Content)');
  });

packagesCmd
  .command('block')
  .description('POST /copilot/admin/catalog/packages/{id}/block')
  .argument('<packageId>', 'Package id')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (packageId: string, opts: AuthOpts, cmd) => {
    checkReadOnly(cmd);
    const token = await resolveTokenOrExit(opts);
    const r = await copilotPackagesBlock(token, packageId);
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    if (r.data !== undefined) printJson(r.data);
    else console.log('OK (204 No Content)');
  });

packagesCmd
  .command('unblock')
  .description('POST /copilot/admin/catalog/packages/{id}/unblock')
  .argument('<packageId>', 'Package id')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (packageId: string, opts: AuthOpts, cmd) => {
    checkReadOnly(cmd);
    const token = await resolveTokenOrExit(opts);
    const r = await copilotPackagesUnblock(token, packageId);
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    if (r.data !== undefined) printJson(r.data);
    else console.log('OK (204 No Content)');
  });

packagesCmd
  .command('reassign')
  .description('POST /copilot/admin/catalog/packages/{id}/reassign')
  .argument('<packageId>', 'Package id')
  .requiredOption('--new-owner-user-id <guid>', 'New owner user id')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (packageId: string, opts: AuthOpts & { newOwnerUserId: string }, cmd) => {
    checkReadOnly(cmd);
    const token = await resolveTokenOrExit(opts);
    const r = await copilotPackagesReassign(token, packageId, opts.newOwnerUserId);
    if (!r.ok) exitGraphError('Error: ', r.error?.message);
    if (r.data !== undefined) printJson(r.data);
    else console.log('OK (204 No Content)');
  });

copilotCommand.addCommand(packagesCmd);

copilotCommand
  .command('notify-help')
  .description('Print Copilot change-notification resource paths (use `subscribe` with --url and a JSON file for encrypted payloads)')
  .action(() => {
    console.log(`Copilot AI interactions (per-user, delegated AiEnterpriseInteraction.Read — include resource data needs encryption):
  /copilot/users/{user-id}/interactionHistory/getAllEnterpriseInteractions
  Optional OData on resource string, e.g. ?$filter=appClass eq 'IPM.SkypeTeams.Message.Copilot.Teams'

Tenant-wide (application AiEnterpriseInteraction.Read.All):
  /copilot/interactionHistory/getAllEnterpriseInteractions

Meeting AI insights are listed per online meeting (GET; OnlineMeetingAiInsight.Read.All):
  /copilot/users/{user-id}/onlineMeetings/{online-meeting-id}/aiInsights
  See: https://learn.microsoft.com/graph/api/onlinemeeting-list-aiinsights

Also: m365-agent-cli subscribe copilot-interactions --user <id> --url <webhook> (shortcut; see subscribe --help)
`);
  });
