/**
 * Graph **beta** — Viva Engage **online meeting** conversation messages, replies, reactions, and linked meeting.
 */

import { callGraphAt, fetchAllPages, GraphApiError, type GraphResponse, graphError } from './graph-client.js';
import { GRAPH_BETA_URL } from './graph-constants.js';

const ROOT = '/communications/onlineMeetingConversations';

function convPath(conversationId: string): string {
  return `${ROOT}/${encodeURIComponent(conversationId.trim())}`;
}

function msgPath(conversationId: string, messageId: string): string {
  return `${convPath(conversationId)}/messages/${encodeURIComponent(messageId.trim())}`;
}

function replyPath(conversationId: string, parentMessageId: string, replyId: string): string {
  return `${msgPath(conversationId, parentMessageId)}/replies/${encodeURIComponent(replyId.trim())}`;
}

async function betaGet(token: string, path: string): Promise<GraphResponse<unknown>> {
  try {
    return await callGraphAt<unknown>(GRAPH_BETA_URL, token, path);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Graph GET failed');
  }
}

async function betaPost(token: string, path: string, body: unknown): Promise<GraphResponse<unknown>> {
  try {
    return await callGraphAt<unknown>(GRAPH_BETA_URL, token, path, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    });
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Graph POST failed');
  }
}

async function betaPatch(
  token: string,
  path: string,
  body: unknown,
  ifMatch?: string
): Promise<GraphResponse<unknown>> {
  const headers: Record<string, string> = { 'Content-Type': 'application/json' };
  if (ifMatch?.trim()) headers['If-Match'] = ifMatch.trim();
  try {
    return await callGraphAt<unknown>(GRAPH_BETA_URL, token, path, {
      method: 'PATCH',
      headers,
      body: JSON.stringify(body)
    });
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Graph PATCH failed');
  }
}

async function betaDelete(token: string, path: string, ifMatch?: string): Promise<GraphResponse<void>> {
  const headers: Record<string, string> = {};
  if (ifMatch?.trim()) headers['If-Match'] = ifMatch.trim();
  try {
    return await callGraphAt<void>(GRAPH_BETA_URL, token, path, { method: 'DELETE', headers }, false);
  } catch (err) {
    if (err instanceof GraphApiError) return graphError(err.message, err.code, err.status);
    return graphError(err instanceof Error ? err.message : 'Graph DELETE failed');
  }
}

function listSuffix(listQuery: string): string {
  if (!listQuery) return '';
  return listQuery.startsWith('?') ? listQuery : `?${listQuery}`;
}

export async function createOnlineMeetingConversationMessage(
  token: string,
  conversationId: string,
  body: unknown
): Promise<GraphResponse<unknown>> {
  return betaPost(token, `${convPath(conversationId)}/messages`, body);
}

export async function patchOnlineMeetingConversationMessage(
  token: string,
  conversationId: string,
  messageId: string,
  body: unknown,
  ifMatch?: string
): Promise<GraphResponse<unknown>> {
  return betaPatch(token, msgPath(conversationId, messageId), body, ifMatch);
}

export async function deleteOnlineMeetingConversationMessage(
  token: string,
  conversationId: string,
  messageId: string,
  ifMatch?: string
): Promise<GraphResponse<void>> {
  return betaDelete(token, msgPath(conversationId, messageId), ifMatch);
}

export async function getOnlineMeetingConversationMessageConversation(
  token: string,
  conversationId: string,
  messageId: string
): Promise<GraphResponse<unknown>> {
  return betaGet(token, `${msgPath(conversationId, messageId)}/conversation`);
}

export async function listOnlineMeetingConversationMessageReactions(
  token: string,
  conversationId: string,
  messageId: string,
  listQuery: string = ''
): Promise<GraphResponse<unknown[]>> {
  return fetchAllPages<unknown>(
    token,
    `${msgPath(conversationId, messageId)}/reactions${listSuffix(listQuery)}`,
    'Failed to list message reactions',
    GRAPH_BETA_URL
  );
}

export async function createOnlineMeetingConversationMessageReaction(
  token: string,
  conversationId: string,
  messageId: string,
  body: unknown
): Promise<GraphResponse<unknown>> {
  return betaPost(token, `${msgPath(conversationId, messageId)}/reactions`, body);
}

export async function getOnlineMeetingConversationMessageReaction(
  token: string,
  conversationId: string,
  messageId: string,
  reactionId: string
): Promise<GraphResponse<unknown>> {
  return betaGet(token, `${msgPath(conversationId, messageId)}/reactions/${encodeURIComponent(reactionId.trim())}`);
}

export async function patchOnlineMeetingConversationMessageReaction(
  token: string,
  conversationId: string,
  messageId: string,
  reactionId: string,
  body: unknown,
  ifMatch?: string
): Promise<GraphResponse<unknown>> {
  return betaPatch(
    token,
    `${msgPath(conversationId, messageId)}/reactions/${encodeURIComponent(reactionId.trim())}`,
    body,
    ifMatch
  );
}

export async function deleteOnlineMeetingConversationMessageReaction(
  token: string,
  conversationId: string,
  messageId: string,
  reactionId: string,
  ifMatch?: string
): Promise<GraphResponse<void>> {
  return betaDelete(
    token,
    `${msgPath(conversationId, messageId)}/reactions/${encodeURIComponent(reactionId.trim())}`,
    ifMatch
  );
}

export async function listOnlineMeetingConversationMessageReplies(
  token: string,
  conversationId: string,
  messageId: string,
  listQuery: string = ''
): Promise<GraphResponse<unknown[]>> {
  return fetchAllPages<unknown>(
    token,
    `${msgPath(conversationId, messageId)}/replies${listSuffix(listQuery)}`,
    'Failed to list message replies',
    GRAPH_BETA_URL
  );
}

export async function createOnlineMeetingConversationMessageReply(
  token: string,
  conversationId: string,
  messageId: string,
  body: unknown
): Promise<GraphResponse<unknown>> {
  return betaPost(token, `${msgPath(conversationId, messageId)}/replies`, body);
}

export async function getOnlineMeetingConversationMessageReply(
  token: string,
  conversationId: string,
  messageId: string,
  replyId: string
): Promise<GraphResponse<unknown>> {
  return betaGet(token, replyPath(conversationId, messageId, replyId));
}

export async function patchOnlineMeetingConversationMessageReply(
  token: string,
  conversationId: string,
  messageId: string,
  replyId: string,
  body: unknown,
  ifMatch?: string
): Promise<GraphResponse<unknown>> {
  return betaPatch(token, replyPath(conversationId, messageId, replyId), body, ifMatch);
}

export async function deleteOnlineMeetingConversationMessageReply(
  token: string,
  conversationId: string,
  messageId: string,
  replyId: string,
  ifMatch?: string
): Promise<GraphResponse<void>> {
  return betaDelete(token, replyPath(conversationId, messageId, replyId), ifMatch);
}

export async function getOnlineMeetingConversationMessageReplyConversation(
  token: string,
  conversationId: string,
  messageId: string,
  replyId: string
): Promise<GraphResponse<unknown>> {
  return betaGet(token, `${replyPath(conversationId, messageId, replyId)}/conversation`);
}

export async function listOnlineMeetingConversationMessageReplyReactions(
  token: string,
  conversationId: string,
  messageId: string,
  replyId: string,
  listQuery: string = ''
): Promise<GraphResponse<unknown[]>> {
  return fetchAllPages<unknown>(
    token,
    `${replyPath(conversationId, messageId, replyId)}/reactions${listSuffix(listQuery)}`,
    'Failed to list reply reactions',
    GRAPH_BETA_URL
  );
}

export async function createOnlineMeetingConversationMessageReplyReaction(
  token: string,
  conversationId: string,
  messageId: string,
  replyId: string,
  body: unknown
): Promise<GraphResponse<unknown>> {
  return betaPost(token, `${replyPath(conversationId, messageId, replyId)}/reactions`, body);
}

export async function getOnlineMeetingConversationMessageReplyReaction(
  token: string,
  conversationId: string,
  messageId: string,
  replyId: string,
  reactionId: string
): Promise<GraphResponse<unknown>> {
  return betaGet(
    token,
    `${replyPath(conversationId, messageId, replyId)}/reactions/${encodeURIComponent(reactionId.trim())}`
  );
}

export async function patchOnlineMeetingConversationMessageReplyReaction(
  token: string,
  conversationId: string,
  messageId: string,
  replyId: string,
  reactionId: string,
  body: unknown,
  ifMatch?: string
): Promise<GraphResponse<unknown>> {
  return betaPatch(
    token,
    `${replyPath(conversationId, messageId, replyId)}/reactions/${encodeURIComponent(reactionId.trim())}`,
    body,
    ifMatch
  );
}

export async function deleteOnlineMeetingConversationMessageReplyReaction(
  token: string,
  conversationId: string,
  messageId: string,
  replyId: string,
  reactionId: string,
  ifMatch?: string
): Promise<GraphResponse<void>> {
  return betaDelete(
    token,
    `${replyPath(conversationId, messageId, replyId)}/reactions/${encodeURIComponent(reactionId.trim())}`,
    ifMatch
  );
}

export async function getOnlineMeetingConversationMessageReplyTo(
  token: string,
  conversationId: string,
  messageId: string
): Promise<GraphResponse<unknown>> {
  return betaGet(token, `${msgPath(conversationId, messageId)}/replyTo`);
}

export async function getOnlineMeetingConversationMessageReplyReplyTo(
  token: string,
  conversationId: string,
  messageId: string,
  replyId: string
): Promise<GraphResponse<unknown>> {
  return betaGet(token, `${replyPath(conversationId, messageId, replyId)}/replyTo`);
}

export async function getOnlineMeetingConversationOnlineMeeting(
  token: string,
  conversationId: string
): Promise<GraphResponse<unknown>> {
  return betaGet(token, `${convPath(conversationId)}/onlineMeeting`);
}
