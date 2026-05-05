import { describe, expect, it } from 'bun:test';
import { mergeExcelSessionInit } from './graph-excel-client.js';

describe('mergeExcelSessionInit', () => {
  it('returns base init unchanged when session id is absent', () => {
    const base = { method: 'PATCH' as const, body: '{}' };
    expect(mergeExcelSessionInit(base, undefined)).toEqual(base);
    expect(mergeExcelSessionInit(base, '')).toEqual(base);
    expect(mergeExcelSessionInit(base, '   ')).toEqual(base);
  });

  it('adds workbook-session-id header', () => {
    const merged = mergeExcelSessionInit({ method: 'POST', body: '{}' }, 'sess-1');
    const h = new Headers(merged.headers as HeadersInit);
    expect(h.get('workbook-session-id')).toBe('sess-1');
  });

  it('preserves existing headers', () => {
    const merged = mergeExcelSessionInit(
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: '{}'
      },
      'abc'
    );
    const h = new Headers(merged.headers as HeadersInit);
    expect(h.get('Content-Type')).toBe('application/json');
    expect(h.get('workbook-session-id')).toBe('abc');
  });

  it('trims session id', () => {
    const merged = mergeExcelSessionInit({ method: 'GET' }, '  trim-me  ');
    expect(new Headers(merged.headers as HeadersInit).get('workbook-session-id')).toBe('trim-me');
  });
});
