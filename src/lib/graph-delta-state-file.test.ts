import { describe, expect, test } from 'bun:test';
import {
  applyDeltaPageToState,
  assertDeltaScopeMatchesState,
  driveDeltaScopeFromLocation,
  parseDeltaStateJson,
  resolveDeltaContinuationUrl
} from './graph-delta-state-file.js';

describe('graph-delta-state-file', () => {
  test('resolveDeltaContinuationUrl prefers explicit next', () => {
    expect(
      resolveDeltaContinuationUrl({
        explicitNext: 'https://graph.microsoft.com/next',
        state: {
          version: 1,
          kind: 'mailMessages',
          updatedAt: '',
          pendingNextLink: 'https://graph.microsoft.com/pending',
          deltaLink: 'https://graph.microsoft.com/delta'
        }
      })
    ).toBe('https://graph.microsoft.com/next');
  });

  test('resolveDeltaContinuationUrl uses pending then delta', () => {
    const s = {
      version: 1 as const,
      kind: 'mailMessages' as const,
      updatedAt: '',
      pendingNextLink: 'https://p',
      deltaLink: 'https://d'
    };
    expect(resolveDeltaContinuationUrl({ state: s })).toBe('https://p');
    expect(
      resolveDeltaContinuationUrl({
        state: { ...s, pendingNextLink: undefined }
      })
    ).toBe('https://d');
  });

  test('applyDeltaPageToState sets deltaLink and clears pending', () => {
    const next = applyDeltaPageToState(null, 'mailMessages', { '@odata.nextLink': 'https://n' }, {});
    expect(next.pendingNextLink).toBe('https://n');
    const done = applyDeltaPageToState(next, 'mailMessages', { '@odata.deltaLink': 'https://dl' }, {});
    expect(done.deltaLink).toBe('https://dl');
    expect(done.pendingNextLink).toBeUndefined();
  });

  test('assertDeltaScopeMatchesState throws on folder mismatch', () => {
    expect(() =>
      assertDeltaScopeMatchesState(
        { version: 1, kind: 'mailMessages', updatedAt: '', folderId: 'a' },
        { folderId: 'b' }
      )
    ).toThrow(/folderId/);
  });

  test('parseDeltaStateJson accepts todoTasks kind', () => {
    const s = parseDeltaStateJson(
      JSON.stringify({
        version: 1,
        kind: 'todoTasks',
        updatedAt: '',
        listId: 'abc'
      })
    );
    expect(s?.kind).toBe('todoTasks');
    expect(s?.listId).toBe('abc');
  });

  test('parseDeltaStateJson accepts driveDelta kind', () => {
    const s = parseDeltaStateJson(
      JSON.stringify({
        version: 1,
        kind: 'driveDelta',
        updatedAt: '',
        driveLocKind: 'me'
      })
    );
    expect(s?.kind).toBe('driveDelta');
    expect(s?.driveLocKind).toBe('me');
  });

  test('assertDeltaScopeMatchesState throws on drive folder mismatch', () => {
    expect(() =>
      assertDeltaScopeMatchesState(
        {
          version: 1,
          kind: 'driveDelta',
          updatedAt: '',
          driveLocKind: 'me',
          driveFolderItemId: 'a'
        },
        driveDeltaScopeFromLocation({ kind: 'me' }, 'b')
      )
    ).toThrow(/driveFolderItemId/);
  });

  test('driveDeltaScopeFromLocation maps site + library', () => {
    expect(driveDeltaScopeFromLocation({ kind: 'site', siteId: 's1', libraryDriveId: 'l1' }, undefined)).toEqual({
      driveLocKind: 'site',
      driveLocSiteId: 's1',
      driveLocLibraryDriveId: 'l1'
    });
  });

  test('parseDeltaStateJson accepts sharePointListItems kind', () => {
    const s = parseDeltaStateJson(
      JSON.stringify({
        version: 1,
        kind: 'sharePointListItems',
        updatedAt: '',
        sharePointSiteId: 'a',
        sharePointListId: 'b'
      })
    );
    expect(s?.kind).toBe('sharePointListItems');
    expect(s?.sharePointSiteId).toBe('a');
  });
});
