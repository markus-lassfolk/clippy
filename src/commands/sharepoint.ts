import { readFile } from 'node:fs/promises';
import { Command } from 'commander';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import {
  applyDeltaPageToState,
  assertDeltaScopeMatchesState,
  readDeltaStateFile,
  resolveDeltaContinuationUrl,
  writeDeltaStateFile
} from '../lib/graph-delta-state-file.js';
import { followSites, listFollowedSites, unfollowSites } from '../lib/graph-insights-client.js';
import {
  createListItem,
  deleteListItem,
  getListItem,
  getListItems,
  getListItemsDeltaPage,
  getLists,
  getSiteByGraphPath,
  getSiteDefaultDriveId,
  updateListItem
} from '../lib/sharepoint-client.js';
import { checkReadOnly } from '../lib/utils.js';

async function parseFieldsJsonOrFile(opts: {
  fields?: string;
  jsonFile?: string;
  mode: 'create' | 'update';
}): Promise<Record<string, any>> {
  let raw: string;
  if (opts.jsonFile?.trim()) {
    try {
      raw = await readFile(opts.jsonFile.trim(), 'utf-8');
    } catch (e) {
      throw new Error(`Could not read --json-file: ${e instanceof Error ? e.message : String(e)}`);
    }
  } else if (opts.fields?.trim()) {
    raw = opts.fields.trim();
  } else {
    throw new Error(opts.mode === 'create' ? 'Provide --fields or --json-file' : 'Provide --fields or --json-file');
  }
  let parsed: unknown;
  try {
    parsed = JSON.parse(raw);
  } catch (e) {
    throw new Error(`Invalid JSON: ${e instanceof Error ? e.message : String(e)}`);
  }
  if (typeof parsed !== 'object' || parsed === null || Array.isArray(parsed)) {
    throw new Error('JSON must be an object');
  }
  const o = parsed as Record<string, unknown>;
  if (opts.mode === 'create') {
    if ('fields' in o && typeof o.fields === 'object' && o.fields !== null && !Array.isArray(o.fields)) {
      return o.fields as Record<string, any>;
    }
    return o as Record<string, any>;
  }
  if ('fields' in o && typeof o.fields === 'object' && o.fields !== null && !Array.isArray(o.fields)) {
    return o.fields as Record<string, any>;
  }
  return o as Record<string, any>;
}

export const sharepointCommand = new Command('sharepoint').description('Manage Microsoft SharePoint Lists').alias('sp');

sharepointCommand
  .command('resolve-site <siteResource>')
  .description(
    'Resolve a site by Graph path (GET /sites/{resource}) — e.g. `contoso.sharepoint.com:/sites/TeamName`. Prints site id and default document library drive id for `files --site-id`.'
  )
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(async (siteResource: string, opts: { json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error || 'Unknown error'}`);
      process.exit(1);
    }
    const site = await getSiteByGraphPath(auth.token, siteResource);
    if (!site.ok || !site.data) {
      console.error(`Error: ${site.error?.message || 'Unknown error'}`);
      process.exit(1);
    }
    const drive = await getSiteDefaultDriveId(auth.token, site.data.id);
    const driveId = drive.ok && drive.data?.id ? drive.data.id : '';
    if (opts.json) {
      console.log(
        JSON.stringify(
          { site: site.data, defaultDriveId: driveId || null, driveError: drive.ok ? null : drive.error },
          null,
          2
        )
      );
      return;
    }
    console.log(`siteId:\t${site.data.id}`);
    if (site.data.displayName) console.log(`name:\t${site.data.displayName}`);
    if (site.data.webUrl) console.log(`webUrl:\t${site.data.webUrl}`);
    if (driveId) {
      console.log(`defaultDriveId:\t${driveId}`);
      console.log(`Example: m365-agent-cli files list --site-id "${site.data.id}"`);
    } else if (!drive.ok) {
      console.error(`Warning: could not load default drive: ${drive.error?.message ?? 'unknown'}`);
    }
  });

sharepointCommand
  .command('lists')
  .description('List all SharePoint lists in a site')
  .requiredOption('--site-id <id>', 'SharePoint Site ID')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(async (opts: { siteId: string; json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error || 'Unknown error'}`);
      process.exit(1);
    }
    const res = await getLists(auth.token, opts.siteId);
    if (!res.ok) {
      console.error(`Error listing lists: ${res.error?.message || 'Unknown error'}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(res.data, null, 2));
      return;
    }
    if (!res.data || res.data.length === 0) {
      console.log('No lists found in this site.');
      return;
    }
    for (const list of res.data) {
      console.log(`${list.name} (${list.id})`);
      if (list.description) console.log(`  ${list.description}`);
    }
  });

sharepointCommand
  .command('items')
  .description('Get items from a SharePoint list')
  .requiredOption('--site-id <id>', 'SharePoint Site ID')
  .requiredOption('--list-id <id>', 'SharePoint List ID')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(async (opts: { siteId: string; listId: string; json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error || 'Unknown error'}`);
      process.exit(1);
    }
    const res = await getListItems(auth.token, opts.siteId, opts.listId);
    if (!res.ok) {
      console.error(`Error getting list items: ${res.error?.message || 'Unknown error'}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(res.data, null, 2));
      return;
    }
    if (!res.data || res.data.length === 0) {
      console.log('No items found in this list.');
      return;
    }
    for (const item of res.data) {
      console.log(`Item ID: ${item.id}`);
      if (item.fields) {
        for (const [key, val] of Object.entries(item.fields)) {
          if (!key.startsWith('@odata')) {
            console.log(`  ${key}: ${val}`);
          }
        }
      }
      console.log('---');
    }
  });

sharepointCommand
  .command('create-item')
  .description(
    'Create an item in a SharePoint list (--fields JSON string or --json-file with field object or { "fields": { … } })'
  )
  .requiredOption('--site-id <id>', 'SharePoint Site ID')
  .requiredOption('--list-id <id>', 'SharePoint List ID')
  .option('--fields <json>', 'JSON string of fields (e.g. \'{"Title": "My Item"}\')')
  .option('--json-file <path>', 'JSON file: either { "fields": { … } } or a flat fields object')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (
      opts: {
        siteId: string;
        listId: string;
        fields?: string;
        jsonFile?: string;
        json?: boolean;
        token?: string;
        identity?: string;
      },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error || 'Unknown error'}`);
        process.exit(1);
      }
      let parsedFields: Record<string, any>;
      try {
        parsedFields = await parseFieldsJsonOrFile({ fields: opts.fields, jsonFile: opts.jsonFile, mode: 'create' });
      } catch (err: any) {
        console.error(`Error: ${err.message}`);
        process.exit(1);
      }
      const res = await createListItem(auth.token, opts.siteId, opts.listId, parsedFields);
      if (!res.ok) {
        console.error(`Error creating list item: ${res.error?.message || 'Unknown error'}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify(res.data, null, 2));
        return;
      }
      console.log(`Successfully created item ${res.data?.id}`);
    }
  );

sharepointCommand
  .command('update-item')
  .description(
    'Update an item in a SharePoint list (--fields or --json-file; file may be { "fields": { … } } or flat fields)'
  )
  .requiredOption('--site-id <id>', 'SharePoint Site ID')
  .requiredOption('--list-id <id>', 'SharePoint List ID')
  .requiredOption('--item-id <id>', 'SharePoint List Item ID')
  .option('--fields <json>', 'JSON string of fields to patch (e.g. \'{"Title": "New Title"}\')')
  .option('--json-file <path>', 'JSON file with fields to PATCH')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (
      opts: {
        siteId: string;
        listId: string;
        itemId: string;
        fields?: string;
        jsonFile?: string;
        json?: boolean;
        token?: string;
        identity?: string;
      },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error || 'Unknown error'}`);
        process.exit(1);
      }
      let parsedFields: Record<string, any>;
      try {
        parsedFields = await parseFieldsJsonOrFile({ fields: opts.fields, jsonFile: opts.jsonFile, mode: 'update' });
      } catch (err: any) {
        console.error(`Error: ${err.message}`);
        process.exit(1);
      }
      const res = await updateListItem(auth.token, opts.siteId, opts.listId, opts.itemId, parsedFields);
      if (!res.ok) {
        console.error(`Error updating list item: ${res.error?.message || 'Unknown error'}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify(res.data, null, 2));
        return;
      }
      console.log(`Successfully updated item ${opts.itemId}`);
    }
  );

sharepointCommand
  .command('get-item')
  .description('Get one SharePoint list item by id ($expand=fields)')
  .requiredOption('--site-id <id>', 'SharePoint Site ID')
  .requiredOption('--list-id <id>', 'SharePoint List ID')
  .requiredOption('--item-id <id>', 'List item ID')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (opts: {
      siteId: string;
      listId: string;
      itemId: string;
      json?: boolean;
      token?: string;
      identity?: string;
    }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error || 'Unknown error'}`);
        process.exit(1);
      }
      const res = await getListItem(auth.token, opts.siteId, opts.listId, opts.itemId);
      if (!res.ok || !res.data) {
        console.error(`Error: ${res.error?.message || 'Unknown error'}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify(res.data, null, 2));
        return;
      }
      console.log(`Item ID: ${res.data.id}`);
      if (res.data.fields) {
        for (const [key, val] of Object.entries(res.data.fields)) {
          if (!key.startsWith('@odata')) {
            console.log(`  ${key}: ${val}`);
          }
        }
      }
    }
  );

sharepointCommand
  .command('delete-item')
  .description('Delete a SharePoint list item')
  .requiredOption('--site-id <id>', 'SharePoint Site ID')
  .requiredOption('--list-id <id>', 'SharePoint List ID')
  .requiredOption('--item-id <id>', 'List item ID')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (
      opts: { siteId: string; listId: string; itemId: string; json?: boolean; token?: string; identity?: string },
      cmd: any
    ) => {
      checkReadOnly(cmd);
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error || 'Unknown error'}`);
        process.exit(1);
      }
      const res = await deleteListItem(auth.token, opts.siteId, opts.listId, opts.itemId);
      if (!res.ok) {
        console.error(`Error: ${res.error?.message || 'Unknown error'}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify({ success: true, itemId: opts.itemId }, null, 2));
        return;
      }
      console.log(`Deleted item ${opts.itemId}`);
    }
  );

sharepointCommand
  .command('items-delta')
  .description(
    'One page of SharePoint list items delta (GET …/items/delta?$expand=fields). Use --url for nextLink/deltaLink; optional --state-file (kind: sharePointListItems).'
  )
  .requiredOption('--site-id <id>', 'SharePoint Site ID')
  .requiredOption('--list-id <id>', 'SharePoint List ID')
  .option('--url <url>', 'Full nextLink or deltaLink URL')
  .option('--state-file <path>', 'Read/write JSON delta cursor')
  .option('--json', 'Output raw page JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(
    async (opts: {
      siteId: string;
      listId: string;
      url?: string;
      stateFile?: string;
      json?: boolean;
      token?: string;
      identity?: string;
    }) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error || 'Unknown error'}`);
        process.exit(1);
      }
      const scope = { sharePointSiteId: opts.siteId.trim(), sharePointListId: opts.listId.trim() };
      const existingState = opts.stateFile ? await readDeltaStateFile(opts.stateFile) : null;
      if (existingState && existingState.kind !== 'sharePointListItems') {
        console.error('Error: state file is not for sharepoint items-delta (kind must be sharePointListItems).');
        process.exit(1);
      }
      try {
        if (existingState) {
          assertDeltaScopeMatchesState(existingState, scope);
        }
      } catch (err) {
        console.error(err instanceof Error ? err.message : err);
        process.exit(1);
      }
      const continueUrl = resolveDeltaContinuationUrl({ explicitNext: opts.url, state: existingState });
      const r = await getListItemsDeltaPage(auth.token, opts.siteId, opts.listId, continueUrl);
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message || 'Delta failed'}`);
        process.exit(1);
      }
      if (opts.stateFile) {
        const merged = applyDeltaPageToState(existingState, 'sharePointListItems', r.data, scope);
        await writeDeltaStateFile(opts.stateFile, merged);
      }
      if (opts.json) {
        console.log(JSON.stringify(r.data, null, 2));
        return;
      }
      console.log(`Changes: ${r.data.value?.length ?? 0} item(s)`);
      if (r.data['@odata.nextLink']) console.log(`nextLink: ${r.data['@odata.nextLink']}`);
      if (r.data['@odata.deltaLink']) console.log(`deltaLink: ${r.data['@odata.deltaLink']}`);
      if (opts.stateFile) console.log(`state-file: ${opts.stateFile} (updated)`);
    }
  );

sharepointCommand
  .command('followed-sites')
  .description('List sites the signed-in user follows (`GET /me/followedSites`).')
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(async (opts: { json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error || 'Unknown error'}`);
      process.exit(1);
    }
    const r = await listFollowedSites(auth.token);
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message || 'followedSites failed'}`);
      process.exit(1);
    }
    const items = r.data.value ?? [];
    if (opts.json) {
      console.log(JSON.stringify({ value: items }, null, 2));
      return;
    }
    if (items.length === 0) {
      console.log('No followed sites.');
      return;
    }
    for (const s of items) {
      const name = s.displayName ?? s.name ?? '(no name)';
      console.log(`${s.id ?? ''}\t${name}`);
      if (s.webUrl) console.log(`  ${s.webUrl}`);
    }
  });

sharepointCommand
  .command('follow <siteId...>')
  .description(
    'Follow one or more SharePoint sites (`POST /me/followedSites/add`). Pass multiple ids to follow many in one call.'
  )
  .option('--json', 'Output as JSON')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(async (siteIds: string[], opts: { json?: boolean; token?: string; identity?: string }, cmd: any) => {
    checkReadOnly(cmd);
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error || 'Unknown error'}`);
      process.exit(1);
    }
    const ids = siteIds.map((s) => s.trim()).filter(Boolean);
    if (ids.length === 0) {
      console.error('Error: provide at least one site id');
      process.exit(1);
    }
    const r = await followSites(auth.token, ids);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message || 'follow failed'}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify(r.data ?? { value: [] }, null, 2));
      return;
    }
    const items = r.data?.value ?? [];
    console.log(`✓ Following ${items.length} site(s)`);
    for (const s of items) {
      const name = s.displayName ?? s.name ?? '(no name)';
      console.log(`  ${s.id ?? ''}\t${name}`);
    }
  });

sharepointCommand
  .command('unfollow <siteId...>')
  .description('Unfollow one or more SharePoint sites (`POST /me/followedSites/remove`).')
  .option('--token <token>', 'Use a specific token')
  .option('--identity <name>', 'Graph token cache identity (default: default)')
  .action(async (siteIds: string[], opts: { token?: string; identity?: string }, cmd: any) => {
    checkReadOnly(cmd);
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error || 'Unknown error'}`);
      process.exit(1);
    }
    const ids = siteIds.map((s) => s.trim()).filter(Boolean);
    if (ids.length === 0) {
      console.error('Error: provide at least one site id');
      process.exit(1);
    }
    const r = await unfollowSites(auth.token, ids);
    if (!r.ok) {
      console.error(`Error: ${r.error?.message || 'unfollow failed'}`);
      process.exit(1);
    }
    console.log(`✓ Unfollowed ${ids.length} site(s)`);
  });
