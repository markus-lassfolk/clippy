import { readFile } from 'node:fs/promises';
import { Command } from 'commander';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import {
  deleteUserItemInsightsSettings,
  deleteWorkingTimeSchedule,
  endWorkingTime,
  getUserItemInsightsSettings,
  getWorkingTimeSchedule,
  listEmployeeExperienceAssignedRoles,
  patchUserItemInsightsSettings,
  patchWorkingTimeSchedule,
  startWorkingTime
} from '../lib/graph-viva-client.js';
import { checkReadOnly } from '../lib/utils.js';

export const vivaCommand = new Command('viva').description(
  'Microsoft Graph **beta** Viva / work-time / insights-related APIs (tenant-dependent; see docs/GRAPH_SCOPES.md)'
);

async function readJsonBody(path: string): Promise<unknown> {
  const raw = await readFile(path.trim(), 'utf8');
  return JSON.parse(raw) as unknown;
}

vivaCommand
  .command('working-time-schedule-get')
  .description('GET /me|/users/{id}/solutions/workingTimeSchedule (beta)')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <id>', 'Target user id or UPN (omit for /me)')
  .action(async (opts: { token?: string; identity?: string; user?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await getWorkingTimeSchedule(auth.token, opts.user);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    console.log(JSON.stringify(r.data ?? null, null, 2));
  });

vivaCommand
  .command('working-time-schedule-patch')
  .description('PATCH workingTimeSchedule (beta); body from --body-file JSON')
  .requiredOption('--body-file <path>', 'JSON object for PATCH body')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <id>', 'Target user id or UPN (omit for /me)')
  .action(async (opts: { bodyFile: string; token?: string; identity?: string; user?: string }, cmd: Command) => {
    checkReadOnly(cmd);
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    let body: unknown;
    try {
      body = await readJsonBody(opts.bodyFile);
    } catch (e) {
      console.error(e instanceof Error ? e.message : 'Invalid --body-file');
      process.exit(1);
    }
    const r = await patchWorkingTimeSchedule(auth.token, body, opts.user);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    console.log(JSON.stringify(r.data ?? null, null, 2));
  });

vivaCommand
  .command('working-time-schedule-delete')
  .description('DELETE workingTimeSchedule (beta); optional If-Match')
  .option('--if-match <etag>', 'If-Match header value')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <id>', 'Target user id or UPN (omit for /me)')
  .action(async (opts: { ifMatch?: string; token?: string; identity?: string; user?: string }, cmd: Command) => {
    checkReadOnly(cmd);
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await deleteWorkingTimeSchedule(auth.token, opts.user, opts.ifMatch);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
  });

vivaCommand
  .command('start-working-time')
  .description('POST .../workingTimeSchedule/startWorkingTime (beta)')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <id>', 'Target user id or UPN (omit for /me)')
  .action(async (opts: { token?: string; identity?: string; user?: string }, cmd: Command) => {
    checkReadOnly(cmd);
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await startWorkingTime(auth.token, opts.user);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
  });

vivaCommand
  .command('end-working-time')
  .description('POST .../workingTimeSchedule/endWorkingTime (beta)')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <id>', 'Target user id or UPN (omit for /me)')
  .action(async (opts: { token?: string; identity?: string; user?: string }, cmd: Command) => {
    checkReadOnly(cmd);
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await endWorkingTime(auth.token, opts.user);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
  });

vivaCommand
  .command('insights-settings-get')
  .description('GET /me|/users/{id}/settings/itemInsights (beta)')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <id>', 'Target user id or UPN (omit for /me)')
  .action(async (opts: { token?: string; identity?: string; user?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await getUserItemInsightsSettings(auth.token, opts.user);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    console.log(JSON.stringify(r.data ?? null, null, 2));
  });

vivaCommand
  .command('insights-settings-patch')
  .description('PATCH user itemInsights / userInsightsSettings (beta)')
  .requiredOption('--body-file <path>', 'JSON object for PATCH body')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <id>', 'Target user id or UPN')
  .action(async (opts: { bodyFile: string; token?: string; identity?: string; user?: string }, cmd: Command) => {
    checkReadOnly(cmd);
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    let body: unknown;
    try {
      body = await readJsonBody(opts.bodyFile);
    } catch (e) {
      console.error(e instanceof Error ? e.message : 'Invalid --body-file');
      process.exit(1);
    }
    const r = await patchUserItemInsightsSettings(auth.token, body, opts.user);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    console.log(JSON.stringify(r.data ?? null, null, 2));
  });

vivaCommand
  .command('insights-settings-delete')
  .description('DELETE /me|/users/{id}/settings/itemInsights (beta)')
  .option('--if-match <etag>', 'If-Match header value')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <id>', 'Target user id or UPN (omit for /me)')
  .action(async (opts: { ifMatch?: string; token?: string; identity?: string; user?: string }, cmd: Command) => {
    checkReadOnly(cmd);
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await deleteUserItemInsightsSettings(auth.token, opts.user, opts.ifMatch);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
  });

vivaCommand
  .command('engage-assigned-roles-list')
  .description('List GET /me|/users/{id}/employeeExperience/assignedRoles (beta, paged)')
  .option('--json', 'Output as JSON array')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <id>', 'Target user id or UPN (omit for /me)')
  .action(async (opts: { json?: boolean; token?: string; identity?: string; user?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await listEmployeeExperienceAssignedRoles(auth.token, opts.user);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message || 'List failed'}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(r.data, null, 2));
    } else {
      for (const row of r.data) {
        console.log(JSON.stringify(row));
      }
    }
  });
