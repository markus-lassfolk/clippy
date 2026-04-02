import { Command } from 'commander';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import {
  createOutlookMasterCategory,
  deleteOutlookMasterCategory,
  isValidOutlookCategoryColor,
  listOutlookMasterCategories,
  OUTLOOK_CATEGORY_COLOR_PRESETS,
  updateOutlookMasterCategory
} from '../lib/outlook-master-categories.js';
import { checkReadOnly } from '../lib/utils.js';

/**
 * Mailbox master category list (display names + preset colors). Same names are used on
 * mail/calendar items when you set categories via EWS (`--category`); To Do uses separate string categories.
 */
export const outlookCategoriesCommand = new Command('outlook-categories').description(
  'List, create, update, or delete Outlook mailbox master categories (names and colors; Graph)'
);

outlookCategoriesCommand
  .command('list')
  .description('Show master categories for the signed-in mailbox')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user (Graph delegation)')
  .action(async (opts: { json?: boolean; token?: string; identity?: string; user?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const result = await listOutlookMasterCategories(auth.token!, opts.user);
    if (!result.ok || !result.data) {
      console.error(`Error: ${result.error?.message}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(result.data, null, 2));
      return;
    }
    if (result.data.length === 0) {
      console.log('No master categories defined.');
      return;
    }
    console.log(`\nOutlook master categories (${result.data.length}):\n`);
    for (const c of result.data) {
      console.log(`  ${c.displayName}`);
      console.log(`    color: ${c.color}   id: ${c.id}`);
    }
    console.log('');
  });

outlookCategoriesCommand
  .command('create')
  .description('Add a category to the mailbox master list (requires MailboxSettings.ReadWrite)')
  .requiredOption('--name <text>', 'Display name (unique in the list)')
  .requiredOption('--color <preset>', `Color: ${OUTLOOK_CATEGORY_COLOR_PRESETS.join(', ')}`)
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user (Graph delegation)')
  .action(
    async (
      opts: { name: string; color: string; json?: boolean; token?: string; identity?: string; user?: string },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const color = opts.color.trim();
      if (!isValidOutlookCategoryColor(color)) {
        console.error(`Invalid --color "${color}". Use one of: ${OUTLOOK_CATEGORY_COLOR_PRESETS.join(', ')}`);
        process.exit(1);
      }
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const result = await createOutlookMasterCategory(auth.token!, opts.name, color, opts.user);
      if (!result.ok || !result.data) {
        console.error(`Error: ${result.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify(result.data, null, 2));
      else {
        console.log(`\nCreated: ${result.data.displayName} (${result.data.color})`);
        console.log(`  id: ${result.data.id}\n`);
      }
    }
  );

outlookCategoriesCommand
  .command('update')
  .description('Rename or recolor a master category by id (from list)')
  .requiredOption('--id <id>', 'Category id')
  .option('--name <text>', 'New display name')
  .option('--color <preset>', `New color: ${OUTLOOK_CATEGORY_COLOR_PRESETS.slice(0, 5).join(', ')}, ... preset24`)
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user (Graph delegation)')
  .action(
    async (
      opts: {
        id: string;
        name?: string;
        color?: string;
        json?: boolean;
        token?: string;
        identity?: string;
        user?: string;
      },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      if (!opts.name && !opts.color) {
        console.error('Error: specify --name and/or --color');
        process.exit(1);
      }
      if (opts.color !== undefined && !isValidOutlookCategoryColor(opts.color.trim())) {
        console.error(`Invalid --color. Use one of: ${OUTLOOK_CATEGORY_COLOR_PRESETS.join(', ')}`);
        process.exit(1);
      }
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const result = await updateOutlookMasterCategory(
        auth.token!,
        opts.id.trim(),
        {
          displayName: opts.name,
          color: opts.color?.trim()
        },
        opts.user
      );
      if (!result.ok || !result.data) {
        console.error(`Error: ${result.error?.message}`);
        process.exit(1);
      }
      if (opts.json) console.log(JSON.stringify(result.data, null, 2));
      else {
        console.log(`\nUpdated: ${result.data.displayName} (${result.data.color})`);
        console.log(`  id: ${result.data.id}\n`);
      }
    }
  );

outlookCategoriesCommand
  .command('delete')
  .description('Remove a category from the master list (does not remove labels from existing items)')
  .requiredOption('--id <id>', 'Category id')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .option('--user <email>', 'Target user (Graph delegation)')
  .action(async (opts: { id: string; json?: boolean; token?: string; identity?: string; user?: string }, cmd: any) => {
    checkReadOnly(cmd);
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const result = await deleteOutlookMasterCategory(auth.token!, opts.id.trim(), opts.user);
    if (!result.ok) {
      console.error(`Error: ${result.error?.message}`);
      process.exit(1);
    }
    if (opts.json) console.log(JSON.stringify({ success: true }, null, 2));
    else console.log(`\nDeleted category id: ${opts.id.trim()}\n`);
  });
