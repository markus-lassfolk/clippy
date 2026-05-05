import { describe, expect, test } from 'bun:test';
import {
  assertCopilotReportPeriod,
  buildCopilotRetrievalBody,
  buildCopilotSearchBody,
  copilotReportPath,
  copilotUserSegment
} from './copilot-graph-client.js';

describe('assertCopilotReportPeriod', () => {
  test('accepts D7', () => {
    expect(assertCopilotReportPeriod('D7')).toBe('D7');
  });
  test('accepts ALL', () => {
    expect(assertCopilotReportPeriod('ALL')).toBe('ALL');
  });
  test('rejects invalid', () => {
    expect(() => assertCopilotReportPeriod('D14')).toThrow(/period must be one of/);
  });
});

describe('copilotReportPath', () => {
  test('builds summary path with OData function and format', () => {
    expect(copilotReportPath('getMicrosoft365CopilotUserCountSummary', 'D30')).toBe(
      "/copilot/reports/getMicrosoft365CopilotUserCountSummary(period='D30')?$format=application/json"
    );
  });
});

describe('buildCopilotSearchBody', () => {
  test('minimal search body', () => {
    expect(buildCopilotSearchBody({ query: 'VPN setup' })).toEqual({ query: 'VPN setup' });
  });
  test('includes pageSize and oneDrive dataSources', () => {
    expect(
      buildCopilotSearchBody({
        query: 'budget',
        pageSize: 10,
        oneDriveFilterExpression: 'path:"https://contoso-my.sharepoint.com/personal/x/Documents/"',
        resourceMetadataNames: ['title', 'author']
      })
    ).toEqual({
      query: 'budget',
      pageSize: 10,
      dataSources: {
        oneDrive: {
          filterExpression: 'path:"https://contoso-my.sharepoint.com/personal/x/Documents/"',
          resourceMetadataNames: ['title', 'author']
        }
      }
    });
  });
  test('rejects empty query', () => {
    expect(() => buildCopilotSearchBody({ query: '   ' })).toThrow(/query is required/);
  });
  test('rejects pageSize out of range', () => {
    expect(() => buildCopilotSearchBody({ query: 'x', pageSize: 0 })).toThrow(/pageSize must be 1–100/);
    expect(() => buildCopilotSearchBody({ query: 'x', pageSize: 101 })).toThrow(/pageSize must be 1–100/);
  });
});

describe('copilotUserSegment', () => {
  test('encodes UPN', () => {
    expect(copilotUserSegment('user@contoso.com')).toBe(encodeURIComponent('user@contoso.com'));
  });
});

describe('buildCopilotRetrievalBody', () => {
  test('builds minimal retrieval body', () => {
    expect(buildCopilotRetrievalBody({ queryString: 'Q1 goals', dataSource: 'sharePoint' })).toEqual({
      queryString: 'Q1 goals',
      dataSource: 'sharePoint'
    });
  });
  test('rejects invalid dataSource', () => {
    expect(() => buildCopilotRetrievalBody({ queryString: 'x', dataSource: 'invalid' })).toThrow(/dataSource must be one of/);
  });
});
