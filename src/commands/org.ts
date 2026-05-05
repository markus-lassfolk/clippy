import { Command } from 'commander';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import {
  getManager,
  getUserProfile,
  listDirectReports,
  listTransitiveReports
} from '../lib/graph-org-client.js';

export const orgCommand = new Command('org').description(
  'Organization directory: user profile, manager, direct reports, transitive reports (Microsoft Graph; see GRAPH_SCOPES.md)'
);

orgCommand
  .command('manager')
  .description('Get manager for /me or another user (GET /me/manager or GET /users/{id}/manager)')
  .option('--user <upn-or-id>', 'User object id or UPN (omit for your own manager)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { user?: string; json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await getManager(auth.token, opts.user);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(r.data, null, 2));
      return;
    }
    const u = r.data;
    console.log(
      `${u.displayName ?? '(no name)'}\t${u.mail ?? u.userPrincipalName ?? ''}\t${u.id}\t${u['@odata.type'] ?? ''}`
    );
  });

orgCommand
  .command('direct-reports')
  .description('List direct reports for /me or another user (GET /me/directReports or GET /users/{id}/directReports)')
  .option('--user <upn-or-id>', 'User object id or UPN (omit for your own reports)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { user?: string; json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await listDirectReports(auth.token, opts.user);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(r.data, null, 2));
      return;
    }
    for (const u of r.data) {
      console.log(`${u.displayName ?? '(no name)'}\t${u.mail ?? u.userPrincipalName ?? ''}\t${u.id}`);
    }
  });

orgCommand
  .command('user')
  .description('Get directory user profile (GET /me or GET /users/{id} with $select)')
  .argument('[upn-or-id]', 'User id or UPN (omit for the signed-in user)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (upnOrId: string | undefined, opts: { json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await getUserProfile(auth.token, upnOrId?.trim() || undefined);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(r.data, null, 2));
      return;
    }
    const u = r.data;
    console.log(`${u.displayName ?? '(no name)'}\t${u.mail ?? u.userPrincipalName ?? ''}\t${u.id}`);
    if (u.jobTitle) console.log(`jobTitle:\t${u.jobTitle}`);
    if (u.department) console.log(`department:\t${u.department}`);
    if (u.officeLocation) console.log(`office:\t${u.officeLocation}`);
    if (u.mobilePhone) console.log(`mobile:\t${u.mobilePhone}`);
    if (u.businessPhones?.length) console.log(`phones:\t${u.businessPhones.join(', ')}`);
  });

orgCommand
  .command('transitive-reports')
  .description('List all reports in the management chain (GET …/transitiveReports)')
  .option('--user <upn-or-id>', 'User object id or UPN (omit for your own subtree)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (opts: { user?: string; json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await listTransitiveReports(auth.token, opts.user);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(r.data, null, 2));
      return;
    }
    for (const u of r.data) {
      console.log(`${u.displayName ?? '(no name)'}\t${u.mail ?? u.userPrincipalName ?? ''}\t${u.id}`);
    }
  });
