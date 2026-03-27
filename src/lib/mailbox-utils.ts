export function resolveMailbox(options: { mailbox?: string }): string | undefined {
  return options.mailbox?.trim() || process.env.EWS_TARGET_MAILBOX?.trim() || undefined;
}
