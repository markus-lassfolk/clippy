import { describe, expect, it } from 'bun:test';

const baseUrl = 'https://graph.microsoft.com/v1.0';

describe('graph-event', () => {
  it('forwardEvent POSTs forward', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async () => new Response(null, { status: 202 })) as unknown as typeof fetch;
      const { forwardEvent } = await import('./graph-event.js');
      const r = await forwardEvent({
        token: 't',
        eventId: 'e1',
        toRecipients: ['a@b.com'],
        comment: 'see below'
      });
      expect(r.ok).toBe(true);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('proposeNewTime POSTs tentativelyAccept with proposedNewTime', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async () => new Response(null, { status: 202 })) as unknown as typeof fetch;
      const { proposeNewTime } = await import('./graph-event.js');
      const r = await proposeNewTime({
        token: 't',
        eventId: 'e1',
        startDateTime: '2026-01-01T10:00:00',
        endDateTime: '2026-01-01T11:00:00',
        timeZone: 'UTC'
      });
      expect(r.ok).toBe(true);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('accept, decline, tentativelyAccept invitations', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async () => new Response(null, { status: 202 })) as unknown as typeof fetch;
      const g = await import('./graph-event.js');
      expect((await g.acceptEventInvitation({ token: 't', eventId: 'e1', comment: 'ok' })).ok).toBe(true);
      expect((await g.declineEventInvitation({ token: 't', eventId: 'e1', sendResponse: false })).ok).toBe(true);
      expect((await g.tentativelyAcceptEventInvitation({ token: 't', eventId: 'e1' })).ok).toBe(true);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});
