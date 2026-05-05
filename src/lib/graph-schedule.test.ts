import { describe, expect, it } from 'bun:test';
import { GraphApiError } from './graph-client.js';

const baseUrl = 'https://graph.microsoft.com/v1.0';
const token = 'tok';

describe('graph-schedule', () => {
  it('getSchedule POSTs calendar/getSchedule with Prefer header', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const urls: string[] = [];
    const bodies: string[] = [];
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async (_input, init) => {
        urls.push(String(_input));
        bodies.push(String(init?.body ?? ''));
        const h = init?.headers;
        const prefer = h instanceof Headers ? h.get('Prefer') : (h as Record<string, string>)?.Prefer;
        expect(prefer).toContain('outlook.timezone');
        return new Response(JSON.stringify({ value: [{ scheduleId: 's1' }] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as typeof fetch;
      const { getSchedule } = await import('./graph-schedule.js');
      const r = await getSchedule(
        token,
        {
          schedules: ['user@x.com'],
          startTime: { dateTime: '2026-05-01T08:00:00', timeZone: 'UTC' },
          endTime: { dateTime: '2026-05-01T18:00:00', timeZone: 'UTC' }
        },
        undefined
      );
      expect(r.ok).toBe(true);
      expect(r.data?.value?.[0]?.scheduleId).toBe('s1');
      expect(urls[0]).toContain('/me/calendar/getSchedule');
      expect(bodies[0]).toContain('schedules');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('getSchedule uses graphUserPath when user set', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const urls: string[] = [];
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async (input) => {
        urls.push(String(input));
        return new Response(JSON.stringify({ value: [] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as typeof fetch;
      const { getSchedule } = await import('./graph-schedule.js');
      await getSchedule(
        token,
        {
          schedules: ['a@b.com'],
          startTime: { dateTime: '2026-05-01T08:00:00', timeZone: 'UTC' },
          endTime: { dateTime: '2026-05-01T09:00:00', timeZone: 'UTC' }
        },
        'delegate@x.com'
      );
      expect(urls[0]).toContain('/users/delegate%40x.com/calendar/getSchedule');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('getSchedule maps GraphApiError from callGraph', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async () => {
        throw new GraphApiError('throttled', 'Throttled', 429);
      }) as typeof fetch;
      const { getSchedule } = await import('./graph-schedule.js');
      const r = await getSchedule(token, {
        schedules: ['x'],
        startTime: { dateTime: '2026-05-01T08:00:00', timeZone: 'UTC' },
        endTime: { dateTime: '2026-05-01T09:00:00', timeZone: 'UTC' }
      });
      expect(r.ok).toBe(false);
      expect(r.error?.message).toContain('throttled');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('getSchedule returns graphError when result not ok', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async () =>
        new Response(JSON.stringify({ error: { message: 'nope' } }), {
          status: 403,
          headers: { 'content-type': 'application/json' }
        })) as typeof fetch;
      const { getSchedule } = await import('./graph-schedule.js');
      const r = await getSchedule(token, {
        schedules: ['x'],
        startTime: { dateTime: '2026-05-01T08:00:00', timeZone: 'UTC' },
        endTime: { dateTime: '2026-05-01T09:00:00', timeZone: 'UTC' }
      });
      expect(r.ok).toBe(false);
      expect(r.error?.message).toContain('nope');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('findMeetingTimes success and error paths', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async () =>
        new Response(JSON.stringify({ meetingTimeSuggestions: [] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        })) as typeof fetch;
      const { findMeetingTimes } = await import('./graph-schedule.js');
      const ok = await findMeetingTimes(token, { timeConstraint: { timeSlots: [] } });
      expect(ok.ok).toBe(true);
    } finally {
      globalThis.fetch = originalFetch;
    }

    try {
      globalThis.fetch = (async () => {
        throw new Error('network');
      }) as typeof fetch;
      const { findMeetingTimes } = await import('./graph-schedule.js');
      const bad = await findMeetingTimes(token, { timeConstraint: { timeSlots: [] } });
      expect(bad.ok).toBe(false);
      expect(bad.error?.message).toContain('network');
    } finally {
      globalThis.fetch = originalFetch;
    }

    try {
      globalThis.fetch = (async () =>
        new Response(JSON.stringify({ error: { message: 'bad' } }), {
          status: 500,
          headers: { 'content-type': 'application/json' }
        })) as typeof fetch;
      const { findMeetingTimes } = await import('./graph-schedule.js');
      const r2 = await findMeetingTimes(token, { timeConstraint: { timeSlots: [] } });
      expect(r2.ok).toBe(false);
      expect(r2.error?.message).toContain('bad');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});
