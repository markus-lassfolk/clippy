/**
 * User-facing hints when calendar event ids don't resolve on Microsoft Graph
 * (mixed Graph vs EWS stacks — ids are not interchangeable).
 */

export const GRAPH_EVENT_ID_HINT =
  'Use an event `id` from the same API as your listing: run `calendar --json` and copy `id` from a row where `backend` matches Graph (`graph` or `auto` with Graph). Microsoft Graph and EWS event ids are not interchangeable.';

export const CALENDAR_EVENT_ID_BACKEND_MISMATCH_HINT =
  'If you copied an id from EWS-style output while using Graph (`M365_EXCHANGE_BACKEND=graph`), switch to `ews`/`auto` for that id or re-list with Graph and use the Graph `id`.';

export interface InvalidGraphEventIdPayload {
  error: string;
  id: string;
  hint: string;
  backendMismatchHint: string;
  graphError?: string;
}

export function buildInvalidGraphEventIdPayload(params: {
  id: string;
  graphGetErrorMessage?: string;
}): InvalidGraphEventIdPayload {
  const detail = params.graphGetErrorMessage?.trim();
  return {
    error: detail
      ? `Invalid or unknown event id: ${params.id} (${detail})`
      : `Invalid or unknown event id: ${params.id}`,
    id: params.id,
    hint: GRAPH_EVENT_ID_HINT,
    backendMismatchHint: CALENDAR_EVENT_ID_BACKEND_MISMATCH_HINT,
    ...(detail ? { graphError: detail } : {})
  };
}
