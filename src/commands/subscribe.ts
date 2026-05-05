import { Command } from 'commander';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import {
  createSubscription,
  deleteSubscription,
  listSubscriptions,
  renewSubscription
} from '../lib/graph-subscriptions.js';
import { getTodoLists } from '../lib/todo-client.js';
import { checkReadOnly } from '../lib/utils.js';

export const subscribeCommand = new Command('subscribe')
  .description(
    'Subscribe to Microsoft Graph push notifications (see also: list, renew, cancel). Resource shortcuts: mail, event, contact, todotask (resolves default To Do list id), copilot-interactions (requires --user). For other resources (e.g. per-meeting Copilot paths), pass the full Graph resource string from Microsoft docs.'
  )
  .argument(
    '[resource]',
    'Resource shortcut: mail, event, contact, todotask, copilot-interactions, or a raw Graph resource path'
  )
  .option('--url <url>', 'Webhook notification URL')
  .option('--expiry <datetime>', 'Expiration datetime (ISO 8601, defaults to 3 days from now)')
  .option('--change-type <type>', 'Change type (comma-separated)', 'created,updated')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Subscribe under this user or shared mailbox (users/{id}/...)')
  .action(async (resource, options, cmd) => {
    if (!resource) {
      return cmd.help();
    }
    if (!options.url) {
      console.error('Error: --url is required.');
      process.exit(1);
    }

    checkReadOnly(cmd);

    // Map friendly resource names to graph endpoints (sync shortcuts only).
    const mapResource = (res: string, user?: string) => {
      const prefix = user?.trim() ? `users/${encodeURIComponent(user.trim())}` : 'me';
      switch (res.toLowerCase()) {
        case 'mail':
          return `${prefix}/messages`;
        case 'event':
          return `${prefix}/events`;
        case 'contact':
          return `${prefix}/contacts`;
        case 'copilot-interactions':
          if (!user?.trim()) {
            console.error('Error: --user is required for copilot-interactions (resource includes user id).');
            process.exit(1);
          }
          return `/copilot/users/${encodeURIComponent(user.trim())}/interactionHistory/getAllEnterpriseInteractions()`;
        default:
          return res;
      }
    };

    async function resolveSubscriptionResource(res: string, user?: string): Promise<string> {
      if (res.toLowerCase() === 'todotask') {
        const auth = await resolveGraphAuth({ token: options.token, identity: options.identity });
        if (!auth.success || !auth.token) {
          console.error(`Error: ${auth.error || 'Graph authentication failed'}`);
          process.exit(1);
        }
        const listsR = await getTodoLists(auth.token, user?.trim() || undefined);
        if (!listsR.ok || !listsR.data?.length) {
          console.error(`Failed to list To Do lists: ${listsR.error?.message || 'unknown error'}`);
          process.exit(1);
        }
        const defaultList = listsR.data.find((l) => l.wellknownListName === 'defaultList');
        if (!defaultList) {
          console.error(
            'Error: No default To Do list found (wellknownListName defaultList). Pass a full resource path, e.g. me/todo/lists/{list-id}/tasks (see `m365-agent-cli todo lists`).'
          );
          process.exit(1);
        }
        const prefix = user?.trim() ? `users/${encodeURIComponent(user.trim())}` : 'me';
        return `${prefix}/todo/lists/${encodeURIComponent(defaultList.id)}/tasks`;
      }
      return mapResource(res, user);
    }

    // Generate clientState for subscription validation (if GRAPH_CLIENT_STATE env is set)
    const clientState = process.env.GRAPH_CLIENT_STATE;

    const graphResource = await resolveSubscriptionResource(resource, options.user);

    // Default expiration to 3 days (Graph allows up to 3 days for most resources)
    let expiry = options.expiry;
    if (!expiry) {
      const date = new Date();
      date.setDate(date.getDate() + 3);
      // Ensure we don't exceed max limits by shaving off a minute
      date.setMinutes(date.getMinutes() - 1);
      expiry = date.toISOString();
    }

    try {
      console.log(`Creating subscription for ${graphResource}...`);
      const res = await createSubscription(
        graphResource,
        options.changeType,
        options.url,
        expiry,
        clientState,
        options.token,
        options.identity
      );
      if (!res.ok) {
        console.error(`Failed to create subscription: ${res.error?.message}`);
        process.exit(1);
      }
      const sub = res.data;
      console.log('Subscription created successfully!');
      console.log(JSON.stringify(sub, null, 2));
    } catch (err) {
      console.error(err instanceof Error ? err.message : err);
      process.exit(1);
    }
  });

subscribeCommand
  .command('list')
  .description('List active subscriptions (Graph GET /subscriptions)')
  .option('--json', 'Output as JSON array')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(async (opts: { json?: boolean; token?: string; identity?: string }) => {
    try {
      const res = await listSubscriptions(opts.token, opts.identity);
      if (!res.ok || !res.data) {
        console.error(`Failed to list subscriptions: ${res.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify(res.data, null, 2));
      else {
        for (const s of res.data) {
          const exp = s.expirationDateTime ?? '';
          console.log(`${s.id}\t${exp}\t${s.resource}`);
        }
      }
    } catch (err) {
      console.error(err instanceof Error ? err.message : err);
      process.exit(1);
    }
  });

subscribeCommand
  .command('renew <id>')
  .description('Extend subscription expiration (Graph PATCH /subscriptions/{id})')
  .requiredOption('--expiry <datetime>', 'New expiration (ISO 8601)')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(async (id: string, opts: { expiry: string; token?: string; identity?: string }, cmd) => {
    try {
      checkReadOnly(cmd);
      const res = await renewSubscription(id, opts.expiry, opts.token, opts.identity);
      if (!res.ok) {
        console.error(`Failed to renew subscription: ${res.error?.message}`);
        process.exit(1);
      }
      console.log('Subscription renewed.');
    } catch (err) {
      console.error(err instanceof Error ? err.message : err);
      process.exit(1);
    }
  });

subscribeCommand
  .command('cancel <id>')
  .description('Cancel an existing subscription')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(async (id, options, cmd) => {
    try {
      checkReadOnly(cmd);
      console.log(`Deleting subscription ${id}...`);
      const res = await deleteSubscription(id, options.token, options.identity);
      if (!res.ok) {
        console.error(`Failed to delete subscription: ${res.error?.message}`);
        process.exit(1);
      }
      console.log('Subscription deleted successfully.');
    } catch (err) {
      console.error(err instanceof Error ? err.message : err);
      process.exit(1);
    }
  });
