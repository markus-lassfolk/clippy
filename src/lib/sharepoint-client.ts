import {
  callGraph,
  callGraphAbsolute,
  fetchAllPages,
  GraphApiError,
  type GraphResponse,
  graphError,
  graphResult
} from './graph-client.js';

export interface SharePointSiteSummary {
  id: string;
  displayName?: string;
  webUrl?: string;
  name?: string;
}

export async function getSiteByGraphPath(
  token: string,
  /** e.g. `contoso.sharepoint.com:/sites/TeamName` (host + `:` + server-relative path) */
  sitePath: string
): Promise<GraphResponse<SharePointSiteSummary>> {
  const encoded = encodeURIComponent(sitePath.trim());
  return callGraph<SharePointSiteSummary>(token, `/sites/${encoded}`);
}

export async function getSiteDefaultDriveId(token: string, siteId: string): Promise<GraphResponse<{ id: string }>> {
  return callGraph<{ id: string }>(token, `/sites/${encodeURIComponent(siteId)}/drive`);
}

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
  let res: GraphResponse<{ value: SharePointList[] }>;
  try {
    res = await callGraph<{ value: SharePointList[] }>(token, `/sites/${siteId}/lists`);
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to get lists');
  }
  if (!res.ok || !res.data?.value) return graphError('Failed to get lists: missing data');
  return graphResult(res.data.value);
}

export async function getListItems(
  token: string,
  siteId: string,
  listId: string
): Promise<GraphResponse<SharePointListItem[]>> {
  return fetchAllPages<SharePointListItem>(
    token,
    `/sites/${siteId}/lists/${encodeURIComponent(listId)}/items?$expand=fields`,
    'Failed to get list items'
  );
}

export async function createListItem(
  token: string,
  siteId: string,
  listId: string,
  fields: Record<string, any>
): Promise<GraphResponse<SharePointListItem>> {
  try {
    return await callGraph<SharePointListItem>(token, `/sites/${siteId}/lists/${encodeURIComponent(listId)}/items`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ fields })
    });
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to create list item');
  }
}

export async function updateListItem(
  token: string,
  siteId: string,
  listId: string,
  itemId: string,
  fields: Record<string, any>
): Promise<GraphResponse<Record<string, any>>> {
  try {
    return await callGraph<Record<string, any>>(
      token,
      `/sites/${siteId}/lists/${encodeURIComponent(listId)}/items/${encodeURIComponent(itemId)}/fields`,
      {
        method: 'PATCH',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(fields)
      }
    );
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to update list item');
  }
}

export async function getListItem(
  token: string,
  siteId: string,
  listId: string,
  itemId: string
): Promise<GraphResponse<SharePointListItem>> {
  try {
    return await callGraph<SharePointListItem>(
      token,
      `/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}/items/${encodeURIComponent(itemId)}?$expand=fields`
    );
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to get list item');
  }
}

export async function deleteListItem(
  token: string,
  siteId: string,
  listId: string,
  itemId: string
): Promise<GraphResponse<void>> {
  try {
    return await callGraph<void>(
      token,
      `/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}/items/${encodeURIComponent(itemId)}`,
      { method: 'DELETE' },
      false
    );
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to delete list item');
  }
}

export interface ListItemsDeltaPage {
  value?: SharePointListItem[];
  '@odata.nextLink'?: string;
  '@odata.deltaLink'?: string;
}

export async function getListItemsDeltaPage(
  token: string,
  siteId: string,
  listId: string,
  nextOrDeltaLink?: string
): Promise<GraphResponse<ListItemsDeltaPage>> {
  try {
    if (nextOrDeltaLink?.trim()) {
      return await callGraphAbsolute<ListItemsDeltaPage>(token, nextOrDeltaLink.trim());
    }
    const path = `/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}/items/delta?$expand=fields`;
    return await callGraph<ListItemsDeltaPage>(token, path);
  } catch (err) {
    if (err instanceof GraphApiError) {
      return graphError(err.message, err.code, err.status);
    }
    return graphError(err instanceof Error ? err.message : 'Failed to fetch list items delta');
  }
}
