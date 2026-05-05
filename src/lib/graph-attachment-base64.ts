/**
 * Strict validation for Graph `fileAttachment.contentBytes` (standard base64).
 * `Buffer.from` does not reject malformed input; callers should use this before POST/upload.
 */
export function decodeGraphFileAttachmentBase64(contentBytes: string): Buffer | null {
  const compact = contentBytes.replace(/\s/g, '');
  if (compact.length === 0) {
    return Buffer.alloc(0);
  }
  if (compact.length % 4 !== 0) {
    return null;
  }
  if (!/^[A-Za-z0-9+/]+={0,2}$/.test(compact)) {
    return null;
  }
  const buf = Buffer.from(compact, 'base64');
  const roundTrip = buf.toString('base64');
  const norm = (s: string) => s.replace(/=+$/, '');
  if (norm(roundTrip) !== norm(compact)) {
    return null;
  }
  return buf;
}
