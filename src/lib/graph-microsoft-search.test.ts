import { describe, expect, test } from 'bun:test';
import { flattenMicrosoftSearchHits, type MicrosoftSearchQueryResponse } from './graph-microsoft-search.js';

describe('flattenMicrosoftSearchHits', () => {
  test('extracts hits from nested value', () => {
    const res: MicrosoftSearchQueryResponse = {
      value: [
        {
          hitsContainers: [
            {
              hits: [
                {
                  rank: 1,
                  hitId: 'h1',
                  summary: 'S',
                  resource: {
                    '@odata.type': '#microsoft.graph.driveItem',
                    id: 'item1',
                    webUrl: 'https://example.com/x',
                    name: 'Doc.docx'
                  }
                }
              ]
            }
          ]
        }
      ]
    };
    const hits = flattenMicrosoftSearchHits(res);
    expect(hits).toHaveLength(1);
    expect(hits[0]).toMatchObject({
      rank: 1,
      hitId: 'h1',
      summary: 'S',
      entityType: 'driveItem',
      id: 'item1',
      webUrl: 'https://example.com/x',
      name: 'Doc.docx'
    });
  });

  test('empty response yields empty array', () => {
    expect(flattenMicrosoftSearchHits({})).toEqual([]);
  });
});
