import { callGraph, GraphResponse } from './graph-client.js';

export interface SitePage {
  id: string;
  name?: string;
  title?: string;
  pageLayout?: string;
  publishingState?: {
    level: string;
    versionId: string;
  };
  webUrl?: string;
  [key: string]: any;
}

export async function listSitePages(token: string, siteId: string): Promise<GraphResponse<{ value: SitePage[] }>> {
  return callGraph<{ value: SitePage[] }>(token, `/sites/${siteId}/pages`);
}

export async function getSitePage(token: string, siteId: string, pageId: string): Promise<GraphResponse<SitePage>> {
  return callGraph<SitePage>(token, `/sites/${siteId}/pages/${pageId}`);
}

export async function updateSitePage(
  token: string,
  siteId: string,
  pageId: string,
  pageData: Partial<SitePage>
): Promise<GraphResponse<SitePage>> {
  return callGraph<SitePage>(
    token,
    `/sites/${siteId}/pages/${pageId}`,
    {
      method: 'PATCH',
      body: JSON.stringify(pageData)
    }
  );
}

export async function publishSitePage(token: string, siteId: string, pageId: string): Promise<GraphResponse<void>> {
  return callGraph<void>(
    token,
    `/sites/${siteId}/pages/${pageId}/publish`,
    {
      method: 'POST'
    },
    false // might not return JSON, just 204 No Content
  );
}
