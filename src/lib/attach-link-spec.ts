/** Parse and validate --attach-link values: "Display name|https://..." or a bare https URL. */

export class AttachmentLinkSpecError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'AttachmentLinkSpecError';
  }
}

function deriveNameFromUrl(urlStr: string): string {
  try {
    const u = new URL(urlStr);
    const path = u.pathname.split('/').filter(Boolean);
    const last = path[path.length - 1];
    if (last && last.length > 0 && last.length < 120) {
      return decodeURIComponent(last.replace(/\+/g, ' '));
    }
    return u.hostname || 'Link';
  } catch {
    return 'Link';
  }
}

/** Only https URLs (no javascript:, file:, etc.). */
export function validateHttpsUrlForAttachment(urlRaw: string): string {
  const trimmed = urlRaw.trim();
  if (!trimmed) {
    throw new AttachmentLinkSpecError('Attachment link URL is empty');
  }
  let url: URL;
  try {
    url = new URL(trimmed);
  } catch {
    throw new AttachmentLinkSpecError(`Invalid attachment link URL: ${trimmed}`);
  }
  if (url.protocol !== 'https:') {
    throw new AttachmentLinkSpecError('Attachment link URL must use https://');
  }
  if (url.username || url.password) {
    throw new AttachmentLinkSpecError('Attachment link URL must not include credentials');
  }
  return url.toString();
}

export interface ParsedAttachLink {
  name: string;
  url: string;
}

export function parseAttachLinkSpec(spec: string): ParsedAttachLink {
  const s = spec.trim();
  if (!s) {
    throw new AttachmentLinkSpecError('Empty --attach-link value');
  }
  const pipe = s.indexOf('|');
  let name: string;
  let urlRaw: string;
  if (pipe === -1) {
    urlRaw = s;
    name = deriveNameFromUrl(urlRaw);
  } else {
    name = s.slice(0, pipe).trim();
    urlRaw = s.slice(pipe + 1).trim();
    if (!name) {
      name = deriveNameFromUrl(urlRaw);
    }
  }
  const url = validateHttpsUrlForAttachment(urlRaw);
  return { name, url };
}
