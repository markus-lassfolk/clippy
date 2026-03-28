import { resolveGraphAuth } from './graph-auth.js';

const GRAPH_BASE_URL = process.env.GRAPH_BASE_URL || 'https://graph.microsoft.com/v1.0';

export interface Subscription {
  id: string;
  resource: string;
  applicationId?: string;
  changeType: string;
  clientState?: string;
  notificationUrl: string;
  expirationDateTime: string;
  creatorId?: string;
}

async function fetchGraph(endpoint: string, options: RequestInit = {}): Promise<Response> {
  const auth = await resolveGraphAuth();
  if (!auth.success || !auth.token) {
    throw new Error(auth.error || 'Failed to authenticate to Graph API');
  }

  const url = `${GRAPH_BASE_URL}${endpoint}`;
  const response = await fetch(url, {
    ...options,
    headers: {
      Authorization: `Bearer ${auth.token}`,
      'Content-Type': 'application/json',
      ...(options.headers || {})
    }
  });

  return response;
}

export async function createSubscription(
  resource: string,
  changeType: string,
  notificationUrl: string,
  expirationDateTime: string
): Promise<Subscription> {
  const response = await fetchGraph('/subscriptions', {
    method: 'POST',
    body: JSON.stringify({
      changeType,
      notificationUrl,
      resource,
      expirationDateTime
    })
  });

  if (!response.ok) {
    const error = await response.text();
    throw new Error(`Failed to create subscription: ${response.status} ${response.statusText} - ${error}`);
  }

  return response.json() as Promise<Subscription>;
}

export async function listSubscriptions(): Promise<Subscription[]> {
  const response = await fetchGraph('/subscriptions', {
    method: 'GET'
  });

  if (!response.ok) {
    const error = await response.text();
    throw new Error(`Failed to list subscriptions: ${response.status} ${response.statusText} - ${error}`);
  }

  const data = (await response.json()) as { value: Subscription[] };
  return data.value;
}

export async function deleteSubscription(id: string): Promise<void> {
  const response = await fetchGraph(`/subscriptions/${id}`, {
    method: 'DELETE'
  });

  if (!response.ok) {
    const error = await response.text();
    throw new Error(`Failed to delete subscription: ${response.status} ${response.statusText} - ${error}`);
  }
}

export async function renewSubscription(id: string, expirationDateTime: string): Promise<void> {
  const response = await fetchGraph(`/subscriptions/${id}`, {
    method: 'PATCH',
    body: JSON.stringify({
      expirationDateTime
    })
  });

  if (!response.ok) {
    const error = await response.text();
    throw new Error(`Failed to renew subscription: ${response.status} ${response.statusText} - ${error}`);
  }
}
