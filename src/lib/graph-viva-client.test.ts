import { describe, expect, test } from 'bun:test';
import { buildVivaListQuery, escapeODataSingleQuotedKey } from './graph-viva-client.js';

describe('graph-viva-client', () => {
  test('escapeODataSingleQuotedKey doubles single quotes', () => {
    expect(escapeODataSingleQuotedKey("a'b")).toBe("a''b");
  });

  test('buildVivaListQuery builds OData query string', () => {
    const q = buildVivaListQuery({
      filter: "displayName eq 'x'",
      select: 'id,displayName',
      top: 10,
      skip: 5,
      count: true
    });
    expect(q.startsWith('?')).toBe(true);
    expect(q).toContain('$filter=');
    expect(q).toContain('$select=');
    expect(q).toContain('$top=10');
    expect(q).toContain('$skip=5');
    expect(q).toContain('$count=true');
  });

  test('buildVivaListQuery returns empty when no options', () => {
    expect(buildVivaListQuery({})).toBe('');
  });
});
