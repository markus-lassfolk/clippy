/**
 * Build Microsoft Graph `sendMail` payload (POST /me/sendMail).
 */

export interface GraphSendFileAttachment {
  name: string;
  contentType: string;
  /** Base64-encoded file content */
  contentBytes: string;
}

export function buildGraphSendMailPayload(opts: {
  to: string[];
  cc?: string[];
  bcc?: string[];
  subject: string;
  body: string;
  html: boolean;
  categories?: string[];
  fileAttachments?: GraphSendFileAttachment[];
}): { message: Record<string, unknown>; saveToSentItems: boolean } {
  const toRecipients = opts.to.map((address) => ({ emailAddress: { address } }));
  const ccRecipients = opts.cc?.filter(Boolean).map((address) => ({ emailAddress: { address } }));
  const bccRecipients = opts.bcc?.filter(Boolean).map((address) => ({ emailAddress: { address } }));

  const body = {
    contentType: opts.html ? 'HTML' : 'Text',
    content: opts.body
  };

  const message: Record<string, unknown> = {
    subject: opts.subject,
    body,
    toRecipients
  };
  if (ccRecipients?.length) message.ccRecipients = ccRecipients;
  if (bccRecipients?.length) message.bccRecipients = bccRecipients;
  if (opts.categories?.length) message.categories = opts.categories;

  if (opts.fileAttachments?.length) {
    message.attachments = opts.fileAttachments.map((a) => ({
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: a.name,
      contentType: a.contentType,
      contentBytes: a.contentBytes
    }));
  }

  return { message, saveToSentItems: true };
}
