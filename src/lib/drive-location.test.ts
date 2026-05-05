import { describe, expect, it } from 'bun:test';
import {
  buildDriveFolderOrRootPath,
  DEFAULT_DRIVE_LOCATION,
  driveDeltaStartPath,
  driveItemPath,
  driveLocationFromCliFlags,
  driveRootPrefix,
  driveRootSearchPath
} from './drive-location.js';

describe('driveRootPrefix', () => {
  it('defaults to me drive', () => {
    expect(driveRootPrefix(DEFAULT_DRIVE_LOCATION)).toBe('/me/drive');
  });

  it('encodes user UPN', () => {
    expect(driveRootPrefix({ kind: 'user', user: 'a@b.com' })).toBe('/users/a%40b.com/drive');
  });

  it('uses explicit drive id', () => {
    expect(driveRootPrefix({ kind: 'drive', driveId: 'abc123' })).toBe('/drives/abc123');
  });

  it('uses site default library', () => {
    expect(driveRootPrefix({ kind: 'site', siteId: 'contoso.sharepoint.com,abc,def' })).toBe(
      '/sites/contoso.sharepoint.com%2Cabc%2Cdef/drive'
    );
  });

  it('uses site plus library drive id', () => {
    expect(driveRootPrefix({ kind: 'site', siteId: 'site1', libraryDriveId: 'libDrive99' })).toBe(
      '/sites/site1/drives/libDrive99'
    );
  });
});

describe('driveLocationFromCliFlags', () => {
  it('rejects conflicting selectors', () => {
    const r = driveLocationFromCliFlags({ user: 'x@y.com', driveId: 'd1' });
    expect(r).toEqual({ error: 'Use only one of --user, --drive-id, or --site-id' });
  });

  it('rejects library without site', () => {
    const r = driveLocationFromCliFlags({ libraryDriveId: 'x' });
    expect(r).toEqual({ error: '--library-drive-id requires --site-id' });
  });

  it('accepts site with library', () => {
    const r = driveLocationFromCliFlags({ siteId: 's1', libraryDriveId: 'l1' });
    expect(r).toEqual({ kind: 'site', siteId: 's1', libraryDriveId: 'l1' });
  });
});

describe('paths', () => {
  it('driveItemPath joins prefix and item', () => {
    expect(driveItemPath({ kind: 'me' }, 'item-1')).toBe('/me/drive/items/item-1');
  });

  it('driveRootSearchPath', () => {
    expect(driveRootSearchPath({ kind: 'me' }, 'foo')).toBe("/me/drive/root/search(q='foo')");
  });

  it('buildDriveFolderOrRootPath root', () => {
    expect(buildDriveFolderOrRootPath({ kind: 'me' })).toBe('/me/drive/root');
  });

  it('buildDriveFolderOrRootPath folder with driveId override', () => {
    expect(buildDriveFolderOrRootPath({ kind: 'me' }, { id: 'f1', driveId: 'd99' })).toBe('/drives/d99/items/f1');
  });

  it('driveDeltaStartPath root', () => {
    expect(driveDeltaStartPath({ kind: 'me' })).toBe('/me/drive/root/delta');
  });

  it('driveDeltaStartPath delegated user folder', () => {
    expect(driveDeltaStartPath({ kind: 'user', user: 'a@b.com' }, 'folder-1')).toBe(
      '/users/a%40b.com/drive/items/folder-1/delta'
    );
  });
});
