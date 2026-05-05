/**
 * Microsoft Graph drive root for OneDrive / SharePoint document libraries.
 * Mutually exclusive: exactly one of me (default), delegated user, explicit drive, or site library.
 */
export type DriveLocation =
  | { kind: 'me' }
  | { kind: 'user'; user: string }
  | { kind: 'drive'; driveId: string }
  | { kind: 'site'; siteId: string; libraryDriveId?: string };

export const DEFAULT_DRIVE_LOCATION: DriveLocation = { kind: 'me' };

export function driveRootPrefix(loc: DriveLocation): string {
  switch (loc.kind) {
    case 'me':
      return '/me/drive';
    case 'user':
      return `/users/${encodeURIComponent(loc.user.trim())}/drive`;
    case 'drive':
      return `/drives/${encodeURIComponent(loc.driveId.trim())}`;
    case 'site': {
      const site = encodeURIComponent(loc.siteId.trim());
      if (loc.libraryDriveId?.trim()) {
        return `/sites/${site}/drives/${encodeURIComponent(loc.libraryDriveId.trim())}`;
      }
      return `/sites/${site}/drive`;
    }
  }
}

/**
 * Path to a folder item or the drive root for colon-path uploads (`…/root:/name:/content`).
 * If `folder.driveId` is set, resolves that drive’s item (cross-drive folder within tenant).
 */
export function buildDriveFolderOrRootPath(loc: DriveLocation, folder?: { id?: string; driveId?: string }): string {
  if (!folder?.id?.trim()) {
    return `${driveRootPrefix(loc)}/root`;
  }
  const id = folder.id.trim();
  if (folder.driveId?.trim()) {
    return `/drives/${encodeURIComponent(folder.driveId.trim())}/items/${encodeURIComponent(id)}`;
  }
  return `${driveRootPrefix(loc)}/items/${encodeURIComponent(id)}`;
}

export function driveItemPath(loc: DriveLocation, itemId: string): string {
  return `${driveRootPrefix(loc)}/items/${encodeURIComponent(itemId.trim())}`;
}

export function driveRootSearchPath(loc: DriveLocation, encodedQuery: string): string {
  return `${driveRootPrefix(loc)}/root/search(q='${encodedQuery}')`;
}

/** First-page path for drive item delta (`…/root/delta` or `…/items/{folderId}/delta`). */
export function driveDeltaStartPath(loc: DriveLocation, folderItemId?: string): string {
  const id = folderItemId?.trim();
  if (id) {
    return `${driveItemPath(loc, id)}/delta`;
  }
  return `${driveRootPrefix(loc)}/root/delta`;
}

export interface DriveLocationCliFlags {
  user?: string;
  driveId?: string;
  siteId?: string;
  libraryDriveId?: string;
  /** Use Microsoft Graph `beta` host for this request (`graph.microsoft.com/beta`). */
  beta?: boolean;
}

export type DriveLocationParseResult = DriveLocation | { error: string };

export function driveLocationFromCliFlags(flags: DriveLocationCliFlags): DriveLocationParseResult {
  const user = flags.user?.trim();
  const driveId = flags.driveId?.trim();
  const siteId = flags.siteId?.trim();
  const libraryDriveId = flags.libraryDriveId?.trim();

  if (libraryDriveId && !siteId) {
    return { error: '--library-drive-id requires --site-id' };
  }

  const modes = [!!user, !!driveId, !!siteId].filter(Boolean).length;
  if (modes > 1) {
    return { error: 'Use only one of --user, --drive-id, or --site-id' };
  }

  if (user) return { kind: 'user', user };
  if (driveId) return { kind: 'drive', driveId };
  if (siteId) return { kind: 'site', siteId, libraryDriveId: libraryDriveId || undefined };
  return DEFAULT_DRIVE_LOCATION;
}
