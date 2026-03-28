import { callGraph, graphResult, graphError } from './graph-client.js';

export type OofStatus = 'alwaysEnabled' | 'scheduled' | 'disabled';

export interface AutomaticRepliesSetting {
  status: OofStatus;
  internalReplyMessage?: string;
  externalReplyMessage?: string;
  scheduledStartDateTime?: string; // ISO 8601
  scheduledEndDateTime?: string;   // ISO 8601
}

export interface MailboxSettings {
  automaticRepliesSetting?: AutomaticRepliesSetting;
}

export interface GetMailboxSettingsResponse {
  automaticRepliesSetting?: AutomaticRepliesSetting;
}

export async function getMailboxSettings(token: string): Promise<{
  ok: boolean;
  data?: GetMailboxSettingsResponse;
  error?: { message: string; code?: string; status?: number };
}> {
  return callGraph<GetMailboxSettingsResponse>(token, '/me/mailboxSettings');
}

export async function setMailboxSettings(
  token: string,
  settings: Partial<AutomaticRepliesSetting>
): Promise<{
  ok: boolean;
  error?: { message: string; code?: string; status?: number };
}> {
  const payload = {
    automaticRepliesSetting: {
      ...(settings.status !== undefined ? { status: settings.status } : {}),
      ...(settings.internalReplyMessage !== undefined
        ? { internalReplyMessage: settings.internalReplyMessage }
        : {}),
      ...(settings.externalReplyMessage !== undefined
        ? { externalReplyMessage: settings.externalReplyMessage }
        : {}),
      ...(settings.scheduledStartDateTime !== undefined
        ? { scheduledStartDateTime: settings.scheduledStartDateTime }
        : {}),
      ...(settings.scheduledEndDateTime !== undefined
        ? { scheduledEndDateTime: settings.scheduledEndDateTime }
        : {})
    }
  };

  const result = await callGraph<Record<string, never>>(
    token,
    '/me/mailboxSettings',
    {
      method: 'PATCH',
      body: JSON.stringify(payload)
    },
    false // don't expect JSON on 204
  );

  if (!result.ok) {
    return {
      ok: false,
      error: result.error || { message: 'Failed to update mailbox settings' }
    };
  }

  return { ok: true };
}
