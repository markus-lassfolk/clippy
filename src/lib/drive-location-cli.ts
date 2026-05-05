import process from 'node:process';
import type { Command } from 'commander';
import { type DriveLocation, type DriveLocationCliFlags, driveLocationFromCliFlags } from './drive-location.js';

const DRIVE_CLI_OPTION_SPECS: readonly [string, string][] = [
  ['--user <upn>', "Target user's default OneDrive (not with --drive-id or --site-id)"],
  ['--drive-id <id>', 'Explicit drive id (e.g. SharePoint document library)'],
  ['--site-id <id>', 'SharePoint site id (default site document library)'],
  ['--library-drive-id <id>', 'Library drive id (only with --site-id)']
];

export function registerDriveLocationCliOptions(cmd: Command): Command {
  for (const [flag, desc] of DRIVE_CLI_OPTION_SPECS) {
    cmd.option(flag, desc);
  }
  return cmd;
}

/** Parse drive CLI flags; on invalid combination prints to stderr and exits with code 1. */
export function resolveDriveLocationForCli(flags: DriveLocationCliFlags): DriveLocation {
  const r = driveLocationFromCliFlags(flags);
  if ('error' in r) {
    console.error(`Error: ${r.error}`);
    process.exit(1);
  }
  return r;
}
