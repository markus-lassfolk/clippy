/**
 * Microsoft Graph path for `mail` when M365_EXCHANGE_BACKEND is graph or auto.
 * Handles list + read-by-id only; returns handled:false for EWS-only options.
 */

import {
  getMessage,
  listAllMailFoldersRecursive,
  listMessagesInFolder,
  type OutlookMessage
} from '../lib/outlook-graph-client.js';

export interface MailGraphCommandOptions {
  limit: string;
  page: string;
  unread?: boolean;
  flagged?: boolean;
  search?: string;
  read?: string;
  download?: string;
  output: string;
  markRead?: string;
  markUnread?: string;
  flag?: string;
  startDate?: string;
  due?: string;
  unflag?: string;
  complete?: string;
  sensitivity?: string;
  move?: string;
  reply?: string;
  replyAll?: string;
  forward?: string;
  to?: string;
  setCategories?: string;
  clearCategories?: string;
  json?: boolean;
  token?: string;
  mailbox?: string;
  identity?: string;
}

function formatDateShort(dateStr: string): string {
  const date = new Date(dateStr);
  const now = new Date();
  const isToday = date.toDateString() === now.toDateString();
  if (isToday) {
    return date.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit', hour12: false });
  }
  return date.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
}

function truncate(str: string, maxLen: number): string {
  if (!str) return '';
  str = str.replace(/\s+/g, ' ').trim();
  if (str.length <= maxLen) return str;
  return `${str.substring(0, maxLen - 1)}\u2026`;
}

const FOLDER_MAP: Record<string, string> = {
  inbox: 'inbox',
  sent: 'sentitems',
  sentitems: 'sentitems',
  drafts: 'drafts',
  deleted: 'deleteditems',
  deleteditems: 'deleteditems',
  trash: 'deleteditems',
  archive: 'archive',
  junk: 'junkemail',
  junkemail: 'junkemail',
  spam: 'junkemail'
};

function graphUnsupported(opts: MailGraphCommandOptions): boolean {
  return Boolean(
    opts.download ||
      opts.search ||
      opts.markRead ||
      opts.markUnread ||
      opts.flag ||
      opts.unflag ||
      opts.complete ||
      opts.sensitivity ||
      opts.move ||
      opts.reply ||
      opts.replyAll ||
      opts.forward ||
      opts.setCategories ||
      opts.clearCategories ||
      opts.startDate ||
      opts.due ||
      opts.flagged
  );
}

/**
 * @returns handled true if the Graph path completed (list or read).
 */
export async function tryMailGraphPortion(
  token: string,
  folderArg: string,
  options: MailGraphCommandOptions,
  _cmd: unknown
): Promise<{ handled: boolean }> {
  if (graphUnsupported(options)) {
    return { handled: false };
  }

  const user = options.mailbox?.trim() || undefined;
  const folderKey = folderArg.toLowerCase();
  let folderId = FOLDER_MAP[folderKey];

  if (!folderId) {
    const all = await listAllMailFoldersRecursive(token, user);
    if (!all.ok || !all.data) {
      console.error(`Error: ${all.error?.message || 'Failed to list folders'}`);
      process.exit(1);
    }
    const found = all.data.find((f) => f.displayName.toLowerCase() === folderArg.toLowerCase());
    if (!found) {
      console.error(`Folder "${folderArg}" not found.`);
      console.error('Use "m365-agent-cli folders" to see available folders.');
      process.exit(1);
    }
    folderId = found.id;
  }

  const limit = Math.max(1, parseInt(options.limit, 10) || 10);
  const page = Math.max(1, parseInt(options.page, 10) || 1);
  const skip = (page - 1) * limit;

  const filters: string[] = [];
  if (options.unread) {
    filters.push('isRead eq false');
  }
  const filter = filters.length ? filters.join(' and ') : undefined;

  if (options.read) {
    const id = options.read.trim();
    const select =
      'subject,body,bodyPreview,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,categories,isRead';
    const full = await getMessage(token, id, user, select);
    if (!full.ok || !full.data) {
      console.error(`Error: ${full.error?.message || 'Failed to fetch email'}`);
      process.exit(1);
    }
    const email = full.data;

    if (options.json) {
      console.log(JSON.stringify(email, null, 2));
      return { handled: true };
    }

    const fromAddr = email.from?.emailAddress?.address ?? '';
    const fromName = email.from?.emailAddress?.name ?? '';
    console.log(`\n${'\u2500'.repeat(60)}`);
    console.log(`From: ${fromName || fromAddr || 'Unknown'}`);
    if (fromAddr) {
      console.log(`      <${fromAddr}>`);
    }
    console.log(`Subject: ${email.subject || '(no subject)'}`);
    const when = email.receivedDateTime || email.sentDateTime;
    console.log(`Date: ${when ? new Date(when).toLocaleString() : 'Unknown'}`);
    if (email.categories?.length) {
      console.log(`Categories: ${email.categories.join(', ')}`);
    }
    console.log(`${'\u2500'.repeat(60)}\n`);
    const content = email.body?.content ?? email.bodyPreview ?? '(no content)';
    console.log(content);
    console.log(`\n${'\u2500'.repeat(60)}\n`);
    return { handled: true };
  }

  const listResult = await listMessagesInFolder(token, folderId, user, {
    top: limit,
    skip,
    orderby: 'receivedDateTime desc',
    filter
  });

  if (!listResult.ok || !listResult.data) {
    console.error(`Error: ${listResult.error?.message || 'Failed to fetch emails'}`);
    process.exit(1);
  }

  const emails = listResult.data;

  if (options.json) {
    console.log(JSON.stringify({ value: emails }, null, 2));
    return { handled: true };
  }

  console.log(`\n${'\u2500'.repeat(60)}`);
  console.log(
    `Folder: ${folderArg} (${emails.length} message${emails.length === 1 ? '' : 's'} shown)${user ? ` — ${user}` : ''}`
  );
  console.log(`${'\u2500'.repeat(60)}\n`);

  if (emails.length === 0) {
    console.log('No messages found.\n');
    return { handled: true };
  }

  for (const m of emails as OutlookMessage[]) {
    const from = m.from?.emailAddress?.address ?? '';
    const subj = truncate(m.subject || '(no subject)', 50);
    const when = m.receivedDateTime ? formatDateShort(m.receivedDateTime) : '';
    const read = m.isRead === false ? ' *' : '';
    console.log(`${when}\t${from}\t${subj}\t${m.id}${read}`);
  }

  console.log(`\n${'\u2500'.repeat(60)}`);
  console.log('\nTip: m365-agent-cli mail -r <id> --read <id>');
  console.log('');
  return { handled: true };
}
