import { validateUrl } from './url-validation';

export const GRAPH_BASE_URL = validateUrl(
  process.env.GRAPH_BASE_URL || 'https://graph.microsoft.com/v1.0',
  'GRAPH_BASE_URL'
);
