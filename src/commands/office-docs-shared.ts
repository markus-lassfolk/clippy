import { readFile } from 'node:fs/promises';
import { basename } from 'node:path';
import type { Command } from 'commander';
import type { DriveLocationCliFlags } from '../lib/drive-location.js';
import { registerDriveLocationCliOptions, resolveDriveLocationForCli } from '../lib/drive-location-cli.js';
import { resolveGraphAuth } from '../lib/graph-auth.js';
import {
  type DriveItem,
  defaultDownloadPath,
  downloadFile,
  getFileMetadata,
  graphApiRoot,
  listDriveItemThumbnails
} from '../lib/graph-client.js';
import { createDriveItemPreview } from '../lib/graph-office-client.js';
import { registerOfficeDriveMirroredCommands } from './office-docs-drive-mirror.js';

type DriveLocOpts = DriveLocationCliFlags;

function graphRoot(flags: { beta?: boolean }): string {
  return graphApiRoot(!!flags.beta);
}

/**
 * Registers `preview`, `meta`, `download`, and `thumbnails` on `word` / `powerpoint` roots (drive-item helpers).
 * Preview: POST …/preview. Meta/download/thumbnails: same patterns as `files` (`drive-location`).
 */
export function registerOfficeDocumentCommands(parent: Command, productLabel: string): void {
  const rootCmd = parent.name();
  parent.description(
    `${productLabel} files on OneDrive / SharePoint drives (Graph). **preview**, **meta**, **download**, **thumbnails** — same as \`files\`. **Mirrored** item APIs (same Graph as \`files\`): **upload**, **upload-large**, **delete**, **share**, **invite**, **permissions**, **permission-remove**, **permission-update**, **copy**, **move**, **versions**, **restore**, **checkout**, **checkin**, **convert**, **analytics**, **activities**, **list-item** (SharePoint columns), **follow**/**unfollow**, **sensitivity-assign**/**sensitivity-extract** (MIP), **retention-label**/**retention-label-remove**, **permanent-delete**. **Folder** **list**/**delta**/**search** → \`files\`. In-file comments / slide OM → not in Graph; **graph invoke** for beta (GRAPH_INVOKE_BOUNDARIES.md).`
  );

  registerDriveLocationCliOptions(
    parent
      .command('preview')
      .description('Create an embeddable preview session for a drive item (POST …/preview)')
      .argument('<itemId>', 'Drive item id')
      .option('--json-file <path>', 'Optional JSON body (e.g. {"chromeless":true}); default {}')
      .option('--json', 'Print raw JSON response')
      .option('--token <token>', 'Graph access token')
      .option('--identity <name>', 'Graph token cache identity')
  ).action(
    async (
      itemId: string,
      opts: DriveLocOpts & { jsonFile?: string; json?: boolean; token?: string; identity?: string }
    ) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      let body: Record<string, unknown> = {};
      if (opts.jsonFile?.trim()) {
        try {
          const raw = await readFile(opts.jsonFile.trim(), 'utf8');
          body = JSON.parse(raw) as Record<string, unknown>;
        } catch (e) {
          console.error(e instanceof Error ? e.message : 'Invalid --json-file');
          process.exit(1);
        }
      }
      const loc = resolveDriveLocationForCli(opts);
      const r = await createDriveItemPreview(auth.token, itemId, body, loc, graphRoot(opts));
      if (!r.ok || !r.data) {
        console.error(`Error: ${r.error?.message}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify(r.data, null, 2));
        return;
      }
      const u = r.data.getUrl ?? r.data.postUrl ?? '';
      console.log(u ? `getUrl/postUrl: ${u}` : JSON.stringify(r.data, null, 2));
    }
  );

  registerDriveLocationCliOptions(
    parent
      .command('meta')
      .description(
        'Get drive item metadata (GET …/drive/items/{id}) — name, size, webUrl, @microsoft.graph.downloadUrl when present'
      )
      .argument('<itemId>', 'Drive item id')
      .option('--json', 'Print full JSON')
      .option('--token <token>', 'Graph access token')
      .option('--identity <name>', 'Graph token cache identity')
  ).action(async (itemId: string, opts: DriveLocOpts & { json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const loc = resolveDriveLocationForCli(opts);
    const r = await getFileMetadata(auth.token, itemId, loc, graphRoot(opts));
    if (!r.ok || !r.data) {
      console.error(`Error: ${r.error?.message}`);
      process.exit(1);
    }
    const d = r.data as DriveItem;
    if (opts.json) {
      console.log(JSON.stringify(d, null, 2));
      return;
    }
    const lines = [
      `name: ${d.name ?? ''}`,
      `id: ${d.id ?? ''}`,
      `size: ${d.size != null ? String(d.size) : ''}`,
      `webUrl: ${d.webUrl ?? ''}`,
      d['@microsoft.graph.downloadUrl']
        ? `downloadUrl: (present — use \`${rootCmd} download\` or \`files download\`)`
        : ''
    ].filter(Boolean);
    console.log(lines.join('\n'));
  });

  registerDriveLocationCliOptions(
    parent
      .command('download')
      .description(
        'Download file bytes by item id (same download path as `files download`; supports drive location flags)'
      )
      .argument('<itemId>', 'Drive item id')
      .option('--out <path>', 'Output path (defaults to ~/Downloads/<name>)')
      .option('--json', 'Output result as JSON (path + item summary)')
      .option('--token <token>', 'Graph access token')
      .option('--identity <name>', 'Graph token cache identity')
  ).action(
    async (
      itemId: string,
      opts: DriveLocOpts & { out?: string; json?: boolean; token?: string; identity?: string }
    ) => {
      const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
      if (!auth.success || !auth.token) {
        console.error(`Auth error: ${auth.error}`);
        process.exit(1);
      }
      const loc = resolveDriveLocationForCli(opts);
      const meta = opts.out ? undefined : await getFileMetadata(auth.token, itemId, loc, graphRoot(opts));
      const defaultOut = meta?.ok && meta.data ? defaultDownloadPath(basename(meta.data.name || itemId)) : undefined;
      const result = await downloadFile(
        auth.token,
        itemId,
        opts.out || defaultOut,
        meta?.ok ? meta.data : undefined,
        loc,
        graphRoot(opts)
      );
      if (!result.ok || !result.data) {
        console.error(`Error: ${result.error?.message || 'Download failed'}`);
        process.exit(1);
      }
      if (opts.json) {
        console.log(JSON.stringify(result.data, null, 2));
        return;
      }
      console.log(`Downloaded: ${result.data.item.name}`);
      console.log(`Saved to: ${result.data.path}`);
    }
  );

  registerDriveLocationCliOptions(
    parent
      .command('thumbnails')
      .description('List thumbnail sets (GET …/thumbnails); same as `files thumbnails`')
      .argument('<itemId>', 'Drive item id')
      .option('--json', 'Output as JSON')
      .option('--token <token>', 'Graph access token')
      .option('--identity <name>', 'Graph token cache identity')
  ).action(async (itemId: string, opts: DriveLocOpts & { json?: boolean; token?: string; identity?: string }) => {
    const auth = await resolveGraphAuth({ token: opts.token, identity: opts.identity });
    if (!auth.success || !auth.token) {
      console.error(`Auth error: ${auth.error}`);
      process.exit(1);
    }
    const loc = resolveDriveLocationForCli(opts);
    const result = await listDriveItemThumbnails(auth.token, itemId, loc, graphRoot(opts));
    if (!result.ok || !result.data) {
      console.error(`Error: ${result.error?.message || 'Failed to list thumbnails'}`);
      process.exit(1);
    }
    if (opts.json) {
      console.log(JSON.stringify({ thumbnails: result.data }, null, 2));
      return;
    }
    if (result.data.length === 0) {
      console.log('No thumbnails returned.');
      return;
    }
    for (const set of result.data) {
      const id = set.id ?? '(set)';
      console.log(`thumbnailSet ${id}`);
      for (const size of ['small', 'medium', 'large'] as const) {
        const info = set[size];
        if (info?.url) {
          console.log(`  ${size}: ${info.width ?? '?'}x${info.height ?? '?'}  ${info.url}`);
        }
      }
    }
  });

  registerOfficeDriveMirroredCommands(parent);
}
