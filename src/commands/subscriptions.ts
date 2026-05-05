import { Command } from 'commander';
import { listSubscriptions, renewSubscription } from '../lib/graph-subscriptions.js';
import { checkReadOnly } from '../lib/utils.js';

export const subscriptionsCommand = new Command('subscriptions').description('Manage Microsoft Graph subscriptions');

function defaultSubscriptionExpiryIso(): string {
  const date = new Date();
  date.setDate(date.getDate() + 3);
  date.setMinutes(date.getMinutes() - 1);
  return date.toISOString();
}

subscriptionsCommand
  .command('renew-all')
  .description(
    'Renew every active subscription to the same expiration (PATCH each id). Default expiry matches `subscribe` (~3 days). For cron: `subscriptions list --json` is optional; this command lists then renews.'
  )
  .option('--expiry <datetime>', 'New expiration for all subscriptions (ISO 8601). Default: ~3 days from now')
  .option('--json', 'Print summary JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(async (opts: { expiry?: string; json?: boolean; token?: string; identity?: string }, cmd) => {
    checkReadOnly(cmd);
    const expiry = opts.expiry?.trim() || defaultSubscriptionExpiryIso();
    const listRes = await listSubscriptions(opts.token, opts.identity);
    if (!listRes.ok || !listRes.data) {
      console.error(`Failed to list subscriptions: ${listRes.error?.message}`);
      process.exit(1);
    }
    const subs = listRes.data;
    const results: Array<{ id: string; ok: boolean; error?: string }> = [];
    for (const s of subs) {
      const r = await renewSubscription(s.id, expiry, opts.token, opts.identity);
      results.push({ id: s.id, ok: r.ok, error: r.ok ? undefined : r.error?.message });
    }
    const failed = results.filter((x) => !x.ok);
    if (opts.json) {
      console.log(
        JSON.stringify(
          {
            expiryDateTime: expiry,
            total: subs.length,
            renewed: results.filter((x) => x.ok).length,
            failed: failed.length,
            results
          },
          null,
          2
        )
      );
    } else {
      console.log(`Renewing ${subs.length} subscription(s) to ${expiry} …`);
      for (const row of results) {
        if (row.ok) console.log(`  ✓ ${row.id}`);
        else console.error(`  ✗ ${row.id}: ${row.error}`);
      }
      if (failed.length > 0) {
        process.exit(1);
      }
      console.log('Done.');
    }
    if (opts.json && failed.length > 0) {
      process.exit(1);
    }
  });

subscriptionsCommand
  .command('list')
  .description('List all active subscriptions')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(async (options: { token?: string; identity?: string }) => {
    try {
      const res = await listSubscriptions(options.token, options.identity);
      if (!res.ok || !res.data) {
        console.error(`Failed to list subscriptions: ${res.error?.message}`);
        process.exit(1);
      }
      const subs = res.data;
      if (subs.length === 0) {
        console.log('No active subscriptions found.');
        return;
      }
      console.log(`Found ${subs.length} active subscription(s):`);
      console.log(JSON.stringify(subs, null, 2));
    } catch (err) {
      console.error(err instanceof Error ? err.message : err);
      process.exit(1);
    }
  });
