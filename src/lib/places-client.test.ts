import { describe, expect, it } from 'bun:test';
import { filterPlacesByQuery } from './places-client.js';

describe('filterPlacesByQuery', () => {
  it('returns all when query blank', () => {
    const places = [{ displayName: 'Room A' }];
    expect(filterPlacesByQuery(places, '   ')).toEqual(places);
  });

  it('matches displayName email building floor tags', () => {
    const places = [
      {
        displayName: 'Alpha',
        emailAddress: 'alpha@x.com',
        building: 'B1',
        floorNumber: '2',
        tags: ['tv', 'whiteboard']
      },
      { displayName: 'Beta', building: 'B2' }
    ];
    expect(filterPlacesByQuery(places, 'tv').map((p) => p.displayName)).toEqual(['Alpha']);
    expect(filterPlacesByQuery(places, 'b1').map((p) => p.displayName)).toEqual(['Alpha']);
    expect(filterPlacesByQuery(places, 'whiteboard').map((p) => p.displayName)).toEqual(['Alpha']);
    expect(filterPlacesByQuery(places, 'beta').map((p) => p.displayName)).toEqual(['Beta']);
  });
});

describe('places-client Graph helpers', () => {
  const token = 'tok';
  const baseUrl = 'https://graph.microsoft.com/v1.0';

  it('getPlace GETs /places/{id}', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const urls: string[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        return new Response(JSON.stringify({ displayName: 'Room X' }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as typeof fetch;

      const { getPlace } = await import('./places-client.js');
      const r = await getPlace(token, 'place-99');
      expect(r.ok).toBe(true);
      expect(r.data?.displayName).toBe('Room X');
      expect(urls[0]).toContain('/places/place-99');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('isRoomFree returns true when calendarView has only free events', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async () =>
        new Response(
          JSON.stringify({
            value: [{ showAs: 'free' }, { showAs: 'free' }]
          }),
          { status: 200, headers: { 'content-type': 'application/json' } }
        )) as typeof fetch;

      const { isRoomFree } = await import('./places-client.js');
      const free = await isRoomFree(token, 'room@contoso.com', '2026-05-01T10:00:00Z', '2026-05-01T11:00:00Z');
      expect(free).toBe(true);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('isRoomFree returns false when busy', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async () =>
        new Response(JSON.stringify({ value: [{ showAs: 'busy' }] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        })) as typeof fetch;

      const { isRoomFree } = await import('./places-client.js');
      const free = await isRoomFree(token, 'room@contoso.com', '2026-05-01T10:00:00Z', '2026-05-01T11:00:00Z');
      expect(free).toBe(false);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('isRoomFree returns null on failed Graph response', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async () =>
        new Response(JSON.stringify({ error: { message: 'x' } }), {
          status: 401,
          headers: { 'content-type': 'application/json' }
        })) as typeof fetch;
      const { isRoomFree } = await import('./places-client.js');
      expect(await isRoomFree(token, 'room@contoso.com', '2026-05-01T10:00:00Z', '2026-05-01T11:00:00Z')).toBeNull();
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('listPlaceRoomLists and listRoomsInRoomList with explicit token', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async (input) => {
        const u = typeof input === 'string' ? input : input instanceof Request ? input.url : String(input);
        if (u.includes('/places/microsoft.graph.roomList')) {
          return new Response(JSON.stringify({ value: [{ id: 'rl1', displayName: 'Building' }] }), {
            status: 200,
            headers: { 'content-type': 'application/json' }
          });
        }
        if (u.includes('lists%40contoso.com')) {
          return new Response(
            JSON.stringify({
              value: [{ displayName: 'R1', capacity: 10, building: 'B', tags: ['tv'] }]
            }),
            { status: 200, headers: { 'content-type': 'application/json' } }
          );
        }
        if (u.includes('/places/microsoft.graph.room')) {
          return new Response(
            JSON.stringify({
              value: [{ displayName: 'R1', capacity: 10, building: 'B', tags: ['tv'] }]
            }),
            { status: 200, headers: { 'content-type': 'application/json' } }
          );
        }
        return new Response(JSON.stringify({ value: [] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as typeof fetch;

      const { listPlaceRoomLists, listRoomsInRoomList, findRooms } = await import('./places-client.js');
      const lists = await listPlaceRoomLists({ token });
      expect(lists.ok).toBe(true);
      expect(lists.data?.[0]?.id).toBe('rl1');
      const rooms = await listRoomsInRoomList('lists@contoso.com', { token });
      expect(rooms.ok).toBe(true);
      expect(rooms.data?.[0]?.displayName).toBe('R1');

      const found = await findRooms({ query: 'r1', building: 'b', capacityMin: 5, equipment: ['tv'] }, { token });
      expect(found.ok).toBe(true);
      expect(found.data?.length).toBe(1);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});
