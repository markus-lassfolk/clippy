import { describe, expect, test } from 'bun:test';
import { buildInvalidGraphEventIdPayload, GRAPH_EVENT_ID_HINT } from './calendar-id-hints.js';

describe('buildInvalidGraphEventIdPayload', () => {
  test('includes graph error and hints', () => {
    const p = buildInvalidGraphEventIdPayload({ id: 'x', graphGetErrorMessage: 'The event is not in the calendar' });
    expect(p.id).toBe('x');
    expect(p.graphError).toBe('The event is not in the calendar');
    expect(p.error).toContain('x');
    expect(p.error).toContain('The event is not in the calendar');
    expect(p.hint).toBe(GRAPH_EVENT_ID_HINT);
    expect(p.backendMismatchHint).toMatch(/EWS/);
  });
});
