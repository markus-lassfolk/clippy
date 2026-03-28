import { callGraph, type GraphResponse } from './graph-client.js';

export interface Recipient {
  emailAddress: {
    address: string;
    name?: string;
  };
}

export interface ForwardEventOptions {
  token: string;
  eventId: string;
  toRecipients: string[];
  comment?: string;
}

export async function forwardEvent(options: ForwardEventOptions): Promise<GraphResponse<void>> {
  const { token, eventId, toRecipients, comment } = options;

  const recipientsList: Recipient[] = toRecipients.map((email) => ({
    emailAddress: { address: email }
  }));

  const body: any = {
    toRecipients: recipientsList
  };

  if (comment) {
    body.comment = comment;
  }

  return callGraph<void>(
    token,
    `/me/events/${encodeURIComponent(eventId)}/forward`,
    {
      method: 'POST',
      body: JSON.stringify(body)
    },
    false
  );
}

export interface ProposeNewTimeOptions {
  token: string;
  eventId: string;
  startDateTime: string;
  endDateTime: string;
  timeZone?: string;
}

export async function proposeNewTime(options: ProposeNewTimeOptions): Promise<GraphResponse<void>> {
  const { token, eventId, startDateTime, endDateTime, timeZone = 'UTC' } = options;

  const body = {
    proposedNewTime: {
      start: { dateTime: startDateTime, timeZone },
      end: { dateTime: endDateTime, timeZone }
    }
  };

  const response = await fetch(`https://graph.microsoft.com/beta/me/events/${encodeURIComponent(eventId)}`, {
    method: 'PATCH',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
      Accept: 'application/json'
    },
    body: JSON.stringify(body)
  });

  if (!response.ok) {
    let message = `Graph request failed: HTTP ${response.status}`;
    let code;
    try {
      const json = await response.json();
      message = json?.error?.message || message;
      code = json?.error?.code;
    } catch {}
    return { ok: false, error: { message, code, status: response.status } };
  }

  return { ok: true, data: undefined };
}
