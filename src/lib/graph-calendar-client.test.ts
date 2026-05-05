import { describe, expect, it } from 'bun:test';

const token = 'tok';
const baseUrl = 'https://graph.microsoft.com/v1.0';

describe('normalizeGraphCalendarRangeInstant', () => {
  it('normalizes date-only and zone-less instants', async () => {
    const { normalizeGraphCalendarRangeInstant } = await import('./graph-calendar-client.js');
    expect(normalizeGraphCalendarRangeInstant('')).toMatch(/1970-01-01/);
    expect(normalizeGraphCalendarRangeInstant('2026-03-15')).toBe('2026-03-15T00:00:00.000Z');
    expect(normalizeGraphCalendarRangeInstant('2026-03-15T14:30:00')).toBe('2026-03-15T14:30:00Z');
    expect(normalizeGraphCalendarRangeInstant('2026-03-15T14:30:00.123')).toBe('2026-03-15T14:30:00Z');
    expect(normalizeGraphCalendarRangeInstant('2026-03-15T14:30:00Z')).toBe('2026-03-15T14:30:00Z');
    expect(normalizeGraphCalendarRangeInstant('2026-03-15T14:30:00+02:00')).toBe('2026-03-15T14:30:00+02:00');
  });
});

describe('graph-calendar-client fetch wrappers', () => {
  it('listCalendarGroups returns pages', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async () =>
        new Response(JSON.stringify({ value: [{ id: 'g1', name: 'G' }] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        })) as typeof fetch;
      const { listCalendarGroups } = await import('./graph-calendar-client.js');
      const r = await listCalendarGroups(token);
      expect(r.ok).toBe(true);
      expect(r.data?.[0]?.id).toBe('g1');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('listCalendars and getCalendar hit user calendar paths', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const urls: string[] = [];
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async (input: string | URL | Request) => {
        const u = typeof input === 'string' ? input : input.toString();
        urls.push(u);
        if (u.includes('/me/calendars/cal-1') && !u.includes('calendarView')) {
          return new Response(JSON.stringify({ id: 'cal-1', name: 'C' }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }
        return new Response(JSON.stringify({ value: [{ id: 'cal-1', name: 'C' }] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as typeof fetch;
      const { listCalendars, getCalendar } = await import('./graph-calendar-client.js');
      const list = await listCalendars(token);
      expect(list.ok).toBe(true);
      expect(urls.some((u) => u.includes('/me/calendars'))).toBe(true);

      const one = await getCalendar(token, 'cal-1');
      expect(one.ok).toBe(true);
      expect(urls.some((u) => u.includes('/me/calendars/cal-1'))).toBe(true);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('createCalendarGroup posts name', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    let posted = '';
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async (_input, init) => {
        posted = String(init?.body ?? '');
        return new Response(JSON.stringify({ id: 'ng', name: 'N' }), {
          status: 201,
          headers: { 'content-type': 'application/json' }
        });
      }) as typeof fetch;
      const { createCalendarGroup } = await import('./graph-calendar-client.js');
      const r = await createCalendarGroup(token, '  N  ');
      expect(r.ok).toBe(true);
      expect(JSON.parse(posted)).toEqual({ name: 'N' });
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('covers calendar events, permissions, attachments, delta, and calendar resource CRUD', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;
    const eventBody = {
      subject: 'S',
      start: { dateTime: '2026-01-01T10:00:00', timeZone: 'UTC' },
      end: { dateTime: '2026-01-01T11:00:00', timeZone: 'UTC' }
    };
    try {
      globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
        const u = typeof input === 'string' ? input : input.toString();
        const m = (init?.method || 'GET').toUpperCase();
        if (m === 'DELETE') {
          return new Response(null, { status: 204 });
        }
        if (m === 'POST' && u.includes('/cancel')) {
          return new Response(null, { status: 202 });
        }
        if (u.includes('/attachments/') && u.includes('/$value')) {
          return new Response(new Uint8Array([7, 8]), { status: 200 });
        }
        if (u.includes('/calendarView')) {
          return new Response(JSON.stringify({ value: [{ id: 'ev-v', subject: 'V' }] }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }
        if (m === 'POST' && u.includes('/events') && !u.includes('/attachments')) {
          return new Response(JSON.stringify({ id: 'new-ev', subject: 'S' }), {
            status: 201,
            headers: { 'content-type': 'application/json' }
          });
        }
        if (m === 'PATCH' && u.includes('/events/')) {
          return new Response(JSON.stringify({ id: 'e1', subject: 'Upd' }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }
        if (u.includes('/instances')) {
          return new Response(JSON.stringify({ value: [{ id: 'inst-1' }] }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }
        if (u.includes('/attachments/') && !u.includes('/$value') && m === 'GET') {
          return new Response(JSON.stringify({ id: 'att1', name: 'a.bin' }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }
        if (u.includes('/attachments') && m === 'GET' && !u.includes('/attachments/')) {
          return new Response(JSON.stringify({ value: [{ id: 'a1' }] }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }
        if (u.includes('/calendarPermissions') && m === 'POST') {
          return new Response(JSON.stringify({ id: 'perm-new', role: 'read' }), {
            status: 201,
            headers: { 'content-type': 'application/json' }
          });
        }
        if (u.includes('/calendarPermissions') && m === 'PATCH') {
          return new Response(JSON.stringify({ id: 'perm-1', role: 'write' }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }
        if (u.includes('/events/delta') && u.includes('$skiptoken=2')) {
          return new Response(
            JSON.stringify({ value: [], '@odata.deltaLink': 'https://graph.microsoft.com/v1.0/me/events/delta' }),
            {
              status: 200,
              headers: { 'content-type': 'application/json' }
            }
          );
        }
        if (u.includes('/events/delta') && m === 'GET') {
          return new Response(
            JSON.stringify({
              value: [{ id: 'd1' }],
              '@odata.nextLink': `${baseUrl}/me/events/delta?$skiptoken=2`
            }),
            { status: 200, headers: { 'content-type': 'application/json' } }
          );
        }
        if (u.includes('/calendarGroups/g1')) {
          return new Response(null, { status: 204 });
        }
        if (m === 'POST' && u.includes('/calendars') && !u.includes('/events')) {
          return new Response(JSON.stringify({ id: 'cal-new', name: 'C2' }), {
            status: 201,
            headers: { 'content-type': 'application/json' }
          });
        }
        if (m === 'PATCH' && u.includes('/calendars/cal-patch')) {
          return new Response(JSON.stringify({ id: 'cal-patch', name: 'Ren' }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }
        return new Response(JSON.stringify({ id: 'e1', subject: 'E' }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as typeof fetch;

      const cal = await import('./graph-calendar-client.js');

      const v = await cal.listCalendarView(token, '2026-01-01T00:00:00Z', '2026-01-02T00:00:00Z');
      expect(v.ok).toBe(true);
      const v2 = await cal.listCalendarView(token, '2026-01-01T00:00:00Z', '2026-01-02T00:00:00Z', {
        calendarId: 'cal-1'
      });
      expect(v2.ok).toBe(true);

      const cr = await cal.createCalendarEvent(token, eventBody);
      expect(cr.ok).toBe(true);
      const cr2 = await cal.createCalendarEvent(token, eventBody, undefined, 'cal-1');
      expect(cr2.ok).toBe(true);

      const up = await cal.updateCalendarEvent(token, 'e1', { subject: 'Upd' });
      expect(up.ok).toBe(true);
      const del = await cal.deleteCalendarEvent(token, 'e1');
      expect(del.ok).toBe(true);
      const can = await cal.cancelCalendarEvent(token, 'e1', { comment: 'x' });
      expect(can.ok).toBe(true);

      const inst = await cal.listEventInstances(token, 'series-1', '2026-01-01', '2026-02-01', {
        select: 'id,subject',
        preferOutlookTimezoneUtc: true
      });
      expect(inst.ok).toBe(true);

      const ge = await cal.getEvent(token, 'e1', undefined, 'id,subject', { preferOutlookTimezoneUtc: true });
      expect(ge.ok).toBe(true);

      const lp = await cal.listCalendarPermissions(token);
      expect(lp.ok).toBe(true);
      const cp = await cal.createCalendarPermission(token, {
        emailAddress: { address: 'x@y.com' },
        role: 'read'
      });
      expect(cp.ok).toBe(true);
      const upp = await cal.updateCalendarPermission(token, 'perm-1', { role: 'write' });
      expect(upp.ok).toBe(true);
      const dp = await cal.deleteCalendarPermission(token, 'perm-1');
      expect(dp.ok).toBe(true);

      const la = await cal.listEventAttachments(token, 'e1');
      expect(la.ok).toBe(true);
      const ga = await cal.getEventAttachment(token, 'e1', 'att1');
      expect(ga.ok).toBe(true);
      const bytes = await cal.downloadEventAttachmentBytes(token, 'e1', 'att1');
      expect(bytes.ok).toBe(true);
      expect(bytes.data?.length).toBe(2);

      const d0 = await cal.eventsDeltaPage(token, {});
      expect(d0.ok).toBe(true);
      const d1 = await cal.eventsDeltaPage(token, {
        nextLink: `${baseUrl}/me/events/delta?$skiptoken=2`
      });
      expect(d1.ok).toBe(true);
      const dCal = await cal.eventsDeltaPage(token, { calendarId: 'cal-1' });
      expect(dCal.ok).toBe(true);

      const dg = await cal.deleteCalendarGroup(token, 'g1');
      expect(dg.ok).toBe(true);

      const cres = await cal.createCalendarResource(token, { name: 'C2', color: 'auto' });
      expect(cres.ok).toBe(true);
      const cgrp = await cal.createCalendarResource(token, { name: 'C3' }, undefined, 'grp-1');
      expect(cgrp.ok).toBe(true);
      const cup = await cal.updateCalendarResource(token, 'cal-patch', { name: 'Ren' });
      expect(cup.ok).toBe(true);
      const cdel = await cal.deleteCalendarResource(token, 'cal-del');
      expect(cdel.ok).toBe(true);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('updateCalendarResource returns error when nothing to patch', async () => {
    const { updateCalendarResource } = await import('./graph-calendar-client.js');
    const r = await updateCalendarResource('tok', 'cal-x', {});
    expect(r.ok).toBe(false);
    expect(r.error?.message).toMatch(/Nothing to update/);
  });
});
