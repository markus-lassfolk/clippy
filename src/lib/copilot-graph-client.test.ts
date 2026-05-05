import { describe, expect, test } from 'bun:test';
import {
  appendCopilotODataQuery,
  assertCopilotReportPeriod,
  buildCopilotRetrievalBody,
  buildCopilotSearchBody,
  COPILOT_CONVERSATION_CHAT_PATH_SUFFIX,
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

describe('appendCopilotODataQuery', () => {
  test('appends query without leading ?', () => {
    expect(appendCopilotODataQuery('/copilot/conversations', '$top=1')).toBe('/copilot/conversations?$top=1');
  });
  test('strips leading ? on fragment', () => {
    expect(appendCopilotODataQuery('/copilot/conversations', '?$top=2')).toBe('/copilot/conversations?$top=2');
  });
  test('merges when path already has query', () => {
    expect(appendCopilotODataQuery('/copilot/conversations?$skip=1', '$top=3')).toBe(
      '/copilot/conversations?$skip=1&$top=3'
    );
  });
});

describe('Copilot chat OData paths', () => {
  test('chat suffix matches Graph OpenAPI action segment', () => {
    expect(COPILOT_CONVERSATION_CHAT_PATH_SUFFIX).toBe('/microsoft.graph.copilot.chat');
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
    expect(() => buildCopilotRetrievalBody({ queryString: 'x', dataSource: 'invalid' })).toThrow(
      /dataSource must be one of/
    );
  });
  test('rejects empty queryString', () => {
    expect(() => buildCopilotRetrievalBody({ queryString: '   ', dataSource: 'sharePoint' })).toThrow(/required/);
  });
  test('rejects queryString over max length', () => {
    const q = 'x'.repeat(1501);
    expect(() => buildCopilotRetrievalBody({ queryString: q, dataSource: 'oneDriveBusiness' })).toThrow(/exceeds/);
  });
  test('includes filterExpression and maximumNumberOfResults', () => {
    expect(
      buildCopilotRetrievalBody({
        queryString: 'q',
        dataSource: 'externalItem',
        filterExpression: '  path:x  ',
        maximumNumberOfResults: 25,
        resourceMetadata: ['a', 'b']
      })
    ).toEqual({
      queryString: 'q',
      dataSource: 'externalItem',
      filterExpression: 'path:x',
      maximumNumberOfResults: 25,
      resourceMetadata: ['a', 'b']
    });
  });
  test('rejects maximumNumberOfResults out of range', () => {
    expect(() =>
      buildCopilotRetrievalBody({ queryString: 'q', dataSource: 'sharePoint', maximumNumberOfResults: 0 })
    ).toThrow(/1–25/);
    expect(() =>
      buildCopilotRetrievalBody({ queryString: 'q', dataSource: 'sharePoint', maximumNumberOfResults: 26 })
    ).toThrow(/1–25/);
  });
});
