import { callGraph, type GraphResponse, graphResult, graphError } from './graph-client.js';

export interface SharePointList {
  id: string;
  name: string;
  displayName: string;
  description?: string;
  createdDateTime: string;
  lastModifiedDateTime: string;
  webUrl: string;
}

export interface SharePointListItem {
  id: string;
  createdDateTime: string;
  lastModifiedDateTime: string;
  webUrl: string;
  fields: Record<string, any>;
}

export async function getLists(token: string, siteId: string): Promise<GraphResponse<SharePointList[]>> {
  const res = await callGraph<{ value: SharePointList[] }>(token, `/sites/${siteId}/lists`);
  if (!res.ok || !res.data) return res as any;
  return graphResult(res.data.value);
}

export async function getListItems(token: string, siteId: string, listId: string): Promise<GraphResponse<SharePointListItem[]>> {
  const res = await callGraph<{ value: SharePointListItem[] }>(token, `/sites/${siteId}/lists/${listId}/items?expand=fields`);
  if (!res.ok || !res.data) return res as any;
  return graphResult(res.data.value);
}

export async function createListItem(
  token: string,
  siteId: string,
  listId: string,
  fields: Record<string, any>
): Promise<GraphResponse<SharePointListItem>> {
  return await callGraph<SharePointListItem>(token, `/sites/${siteId}/lists/${listId}/items`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({ fields })
  });
}

export async function updateListItem(
  token: string,
  siteId: string,
  listId: string,
  itemId: string,
  fields: Record<string, any>
): Promise<GraphResponse<SharePointListItem>> {
  return await callGraph<SharePointListItem>(token, `/sites/${siteId}/lists/${listId}/items/${itemId}/fields`, {
    method: 'PATCH',
    headers: {
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(fields)
  });
}
