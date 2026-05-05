/**
 * Outlook contact merge suggestions visibility (Microsoft Graph **beta** userSettings).
 * @see https://learn.microsoft.com/graph/api/resources/contactmergesuggestions
 */
import {
  callGraphAt,
  GraphApiError,
  type GraphResponse,
  graphError,
  graphErrorFromApiError,
  graphResult
} from './graph-client.js';
import { getGraphBetaUrl } from './graph-constants.js';
import { graphUserPath } from './graph-user-path.js';

export type ContactMergeSuggestionsJson = Record<string, unknown>;

function mergeSuggestionsPath(user?: string): string {
  return graphUserPath(user, 'settings/contactMergeSuggestions');
}

/** `GET …/settings/contactMergeSuggestions` (beta). */
export async function getContactMergeSuggestions(
  token: string,
  user?: string
): Promise<GraphResponse<ContactMergeSuggestionsJson>> {
  try {
    return await callGraphAt<ContactMergeSuggestionsJson>(
      getGraphBetaUrl(),
      token,
      mergeSuggestionsPath(user),
      { method: 'GET' },
      true
    );
  } catch (err) {
    if (err instanceof GraphApiError) return graphErrorFromApiError(err);
    return graphError(err instanceof Error ? err.message : 'Failed to get contactMergeSuggestions');
  }
}

/** `PATCH …/settings/contactMergeSuggestions` (beta). */
export async function patchContactMergeSuggestions(
  token: string,
  body: ContactMergeSuggestionsJson,
  user?: string
): Promise<GraphResponse<ContactMergeSuggestionsJson>> {
  try {
    const r = await callGraphAt<ContactMergeSuggestionsJson>(
      getGraphBetaUrl(),
      token,
      mergeSuggestionsPath(user),
      { method: 'PATCH', body: JSON.stringify(body) },
      true
    );
    if (!r.ok || !r.data) {
      return graphError(
        r.error?.message || 'Failed to patch contactMergeSuggestions',
        r.error?.code,
        r.error?.status
      );
    }
    return graphResult(r.data);
  } catch (err) {
    if (err instanceof GraphApiError) return graphErrorFromApiError(err);
    return graphError(err instanceof Error ? err.message : 'Failed to patch contactMergeSuggestions');
  }
}

/** `DELETE …/settings/contactMergeSuggestions` (beta) — requires If-Match on the resource. */
export async function deleteContactMergeSuggestions(
  token: string,
  ifMatch: string,
  user?: string
): Promise<GraphResponse<void>> {
  try {
    const r = await callGraphAt<void>(
      getGraphBetaUrl(),
      token,
      mergeSuggestionsPath(user),
      { method: 'DELETE', headers: { 'If-Match': ifMatch.trim() } },
      false
    );
    if (!r.ok) {
      return graphError(
        r.error?.message || 'Failed to delete contactMergeSuggestions',
        r.error?.code,
        r.error?.status
      );
    }
    return graphResult(undefined as undefined);
  } catch (err) {
    if (err instanceof GraphApiError) return graphErrorFromApiError(err);
    return graphError(err instanceof Error ? err.message : 'Failed to delete contactMergeSuggestions');
  }
}
