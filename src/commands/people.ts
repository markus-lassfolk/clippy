import { Command } from 'commander';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import { getPerson, listPeople, type Person } from '../lib/graph-directory.js';

function printPersonBrief(p: Person): void {
  const email = p.userPrincipalName || p.scoredEmailAddresses?.[0]?.address || '';
  console.log(`${p.displayName ?? '(no name)'}\t${email}\t${p.id}`);
  if (p.jobTitle) console.log(`  jobTitle: ${p.jobTitle}`);
  if (p.department) console.log(`  department: ${p.department}`);
}

export const peopleCommand = new Command('people').description(
  'Microsoft Graph People API: relevant contacts for /me or another user (People.Read / People.Read.All; see GRAPH_SCOPES.md)'
);

peopleCommand
  .command('list')
  .description('List people ordered by relevance (GET /me/people or GET /users/{id}/people)')
  .option('--user <upn-or-id>', 'Whose people list (omit for the signed-in user)')
  .option('--search <text>', 'Graph $search (requires ConsistencyLevel; same style as `find`)')
  .option('--top <n>', 'Return only the first page of up to n items (omit to follow all @odata.nextLink pages)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(
    async (opts: {
      user?: string;
      search?: string;
      top?: string;
      json?: boolean;
      token?: string;
      identity?: string;
    }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      let top: number | undefined;
      if (opts.top !== undefined) {
        top = parseInt(opts.top, 10);
        if (Number.isNaN(top) || top < 1) {
          console.error('Error: --top must be a positive integer.');
          process.exit(1);
        }
      }
      const r = await listPeople(auth.token, {
        user: opts.user,
        search: opts.search,
        top
      });
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify({ people: r.data }, null, 2));
        return;
      }
      if (r.data.length === 0) {
        console.log('No people returned.');
        return;
      }
      for (const p of r.data) {
        printPersonBrief(p);
        console.log('');
      }
    }
  );

peopleCommand
  .command('get')
  .description('Get one person by id (GET /me/people/{id} or GET /users/{user}/people/{id})')
  .argument('<person-id>', 'Person id from list or search')
  .option('--user <upn-or-id>', 'User whose people graph to read (omit for /me/people)')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Graph access token')
  .option('--identity <name>', 'Graph token cache identity')
  .action(async (personId: string, opts: { user?: string; json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const r = await getPerson(auth.token, personId, opts.user);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(r.data, null, 2));
      return;
    }
    printPersonBrief(r.data);
  });
