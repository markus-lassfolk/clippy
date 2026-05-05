import { describe, expect, test } from 'bun:test';
import { decodeGraphFileAttachmentBase64 } from './graph-attachment-base64.js';

describe('decodeGraphFileAttachmentBase64', () => {
  test('decodes valid standard base64', () => {
    const buf = decodeGraphFileAttachmentBase64('Zm9v'); // "foo"
    expect(buf).not.toBeNull();
    expect(buf!.toString()).toBe('foo');
  });

  test('ignores whitespace', () => {
    const buf = decodeGraphFileAttachmentBase64(' Zm9v \n');
    expect(buf).not.toBeNull();
    expect(buf!.toString()).toBe('foo');
  });

  test('empty and whitespace-only yield empty buffer', () => {
    expect(decodeGraphFileAttachmentBase64('')!.length).toBe(0);
    expect(decodeGraphFileAttachmentBase64('  \t')!.length).toBe(0);
  });

  test('rejects non-base64 characters', () => {
    expect(decodeGraphFileAttachmentBase64('@@@@')).toBeNull();
  });

  test('rejects wrong length (not multiple of 4)', () => {
    expect(decodeGraphFileAttachmentBase64('abc')).toBeNull();
  });

  test('rejects padding that does not round-trip', () => {
    expect(decodeGraphFileAttachmentBase64('Zm9v====')).toBeNull();
  });
});
