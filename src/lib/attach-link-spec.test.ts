import { describe, expect, test } from 'bun:test';
import { AttachmentLinkSpecError, parseAttachLinkSpec, validateHttpsUrlForAttachment } from './attach-link-spec.js';

describe('attach-link-spec', () => {
  test('validateHttpsUrlForAttachment accepts https', () => {
    expect(validateHttpsUrlForAttachment('https://example.com/path?q=1')).toBe('https://example.com/path?q=1');
  });

  test('validateHttpsUrlForAttachment rejects http', () => {
    expect(() => validateHttpsUrlForAttachment('http://example.com')).toThrow(AttachmentLinkSpecError);
  });

  test('parseAttachLinkSpec splits name and url', () => {
    const p = parseAttachLinkSpec('Agenda|https://example.com/doc.pdf');
    expect(p.name).toBe('Agenda');
    expect(p.url).toBe('https://example.com/doc.pdf');
  });

  test('parseAttachLinkSpec bare url derives name', () => {
    const p = parseAttachLinkSpec('https://example.com/folder/doc.pdf');
    expect(p.url).toBe('https://example.com/folder/doc.pdf');
    expect(p.name).toBe('doc.pdf');
  });
});
