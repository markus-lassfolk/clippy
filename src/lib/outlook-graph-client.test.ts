import { afterEach, describe, expect, it } from 'bun:test';

const token = 'test-token';
const baseUrl = 'https://graph.microsoft.com/v1.0';

describe('listMailFolders', () => {
  it('GETs /mailFolders collection', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const urls: string[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        return new Response(
          JSON.stringify({
            value: [{ id: 'inbox-id', displayName: 'Inbox' }]
          }),
          { status: 200, headers: { 'content-type': 'application/json' } }
        );
      }) as typeof fetch;

      const { listMailFolders } = await import('./outlook-graph-client.js');
      const r = await listMailFolders(token);

      expect(r.ok).toBe(true);
      expect(r.data?.[0]?.displayName).toBe('Inbox');
      expect(urls[0]).toContain('/me/mailFolders');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});

describe('getMessage', () => {
  it('GETs /messages/{id}', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const urls: string[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        return new Response(JSON.stringify({ id: 'msg-1', subject: 'Hi', isRead: false }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as typeof fetch;

      const { getMessage } = await import('./outlook-graph-client.js');
      const r = await getMessage(token, 'msg-1', undefined, 'subject,isRead');

      expect(r.ok).toBe(true);
      expect(r.data?.subject).toBe('Hi');
      expect(urls[0]).toContain('/me/messages/msg-1');
      expect(urls[0]).toContain('$select=');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});

describe('listMailboxMessages', () => {
  it('GETs /me/messages', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const urls: string[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        return new Response(JSON.stringify({ value: [{ id: 'm1', subject: 'A' }] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as typeof fetch;

      const { listMailboxMessages } = await import('./outlook-graph-client.js');
      const r = await listMailboxMessages(token, undefined, { top: 10 });

      expect(r.ok).toBe(true);
      expect(urls[0]).toContain('/me/messages');
      expect(decodeURIComponent(urls[0])).toContain('$top=10');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('adds ConsistencyLevel and quoted $search when using search', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    let consistency: string | undefined;
    let requestUrl = '';
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
        requestUrl = typeof input === 'string' ? input : input.toString();
        const h = init?.headers;
        if (h instanceof Headers) {
          consistency = h.get('ConsistencyLevel') ?? undefined;
        } else if (h && typeof h === 'object') {
          consistency = (h as Record<string, string>).ConsistencyLevel;
        }
        return new Response(JSON.stringify({ value: [] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as typeof fetch;

      const { listMailboxMessages } = await import('./outlook-graph-client.js');
      const r = await listMailboxMessages(token, undefined, { top: 5, search: 'budget' });

      expect(r.ok).toBe(true);
      expect(consistency).toBe('eventual');
      expect(decodeURIComponent(requestUrl)).toContain('$search="budget"');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});

describe('sendMail', () => {
  it('POSTs /sendMail', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const urls: string[] = [];
    const bodies: string[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        if (init?.body && typeof init.body === 'string') bodies.push(init.body);
        return new Response(null, { status: 202 });
      }) as typeof fetch;

      const { sendMail } = await import('./outlook-graph-client.js');
      const r = await sendMail(token, {
        message: { subject: 'Hi', body: { contentType: 'Text', content: 'x' } },
        saveToSentItems: true
      });

      expect(r.ok).toBe(true);
      expect(urls[0]).toContain('/me/sendMail');
      expect(bodies[0]).toContain('saveToSentItems');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});

describe('createDraftMessage', () => {
  it('POSTs /me/messages with isDraft', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const urls: string[] = [];
    const bodies: string[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        if (init?.body && typeof init.body === 'string') bodies.push(init.body);
        return new Response(JSON.stringify({ id: 'draft-1', isDraft: true }), {
          status: 201,
          headers: { 'content-type': 'application/json' }
        });
      }) as typeof fetch;

      const { createDraftMessage } = await import('./outlook-graph-client.js');
      const r = await createDraftMessage(token, {
        subject: 'S',
        bodyContent: 'hello',
        bodyContentType: 'Text',
        toAddresses: ['a@b.com']
      });

      expect(r.ok).toBe(true);
      expect(r.data?.id).toBe('draft-1');
      expect(urls[0]).toContain('/me/messages');
      expect(bodies[0]).toContain('"isDraft":true');
      expect(bodies[0]).toContain('a@b.com');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});

describe('mailMessagesDeltaPage', () => {
  it('GETs /me/messages/delta when no folder', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const urls: string[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        return new Response(JSON.stringify({ value: [{ id: 'm1' }] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as typeof fetch;

      const { mailMessagesDeltaPage } = await import('./outlook-graph-client.js');
      const r = await mailMessagesDeltaPage(token, {});

      expect(r.ok).toBe(true);
      expect(r.data?.value?.[0]?.id).toBe('m1');
      expect(urls[0]).toContain('/me/messages/delta');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('GETs .../mailFolders/{id}/messages/delta when folder set', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const urls: string[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        return new Response(JSON.stringify({ value: [] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as typeof fetch;

      const { mailMessagesDeltaPage } = await import('./outlook-graph-client.js');
      const r = await mailMessagesDeltaPage(token, { folderId: 'inbox' });

      expect(r.ok).toBe(true);
      expect(urls[0]).toContain('/me/mailFolders/inbox/messages/delta');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});

describe('mail folders for delegated user', () => {
  it('listMailFolders uses /users/{upn}/mailFolders', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const urls: string[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        return new Response(JSON.stringify({ value: [{ id: 'inbox', displayName: 'Inbox' }] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as typeof fetch;

      const { listMailFolders } = await import('./outlook-graph-client.js');
      const r = await listMailFolders(token, 'shared@contoso.com');

      expect(r.ok).toBe(true);
      expect(urls[0]).toContain('/users/shared%40contoso.com/mailFolders');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('getMailFolder GETs folder by id', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const urls: string[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        return new Response(JSON.stringify({ id: 'fld-1', displayName: 'Archive' }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as typeof fetch;

      const { getMailFolder } = await import('./outlook-graph-client.js');
      const r = await getMailFolder(token, 'fld-1', 'u@contoso.com');

      expect(r.ok).toBe(true);
      expect(r.data?.displayName).toBe('Archive');
      expect(urls[0]).toContain('/users/u%40contoso.com/mailFolders/fld-1');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('listMailboxMessages uses delegated path when user set', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const urls: string[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        return new Response(JSON.stringify({ value: [{ id: 'm1' }] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as typeof fetch;

      const { listMailboxMessages } = await import('./outlook-graph-client.js');
      const r = await listMailboxMessages(token, 'delegate@contoso.com', { top: 3 });

      expect(r.ok).toBe(true);
      expect(urls[0]).toContain(`/users/${encodeURIComponent('delegate@contoso.com')}/messages`);
      expect(urls[0]).toMatch(/[?&](\$|%24)top=3/);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});

describe('listContacts structured query', () => {
  it('GETs /me/contacts with $orderby when using ContactListQueryOptions', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const urls: string[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        return new Response(JSON.stringify({ value: [{ id: 'c1', displayName: 'A' }] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as typeof fetch;

      const { listContacts } = await import('./outlook-graph-client.js');
      const r = await listContacts(token, undefined, { orderby: 'displayName asc' });

      expect(r.ok).toBe(true);
      expect(r.data?.[0]?.displayName).toBe('A');
      expect(decodeURIComponent(urls[0])).toContain('/me/contacts');
      expect(decodeURIComponent(urls[0])).toContain('$orderby=displayName+asc');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});

describe('contact open extensions (folder paths)', () => {
  it('listContactOpenExtensions uses contactFolders path when location.folderId set', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const urls: string[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        return new Response(JSON.stringify({ value: [] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as typeof fetch;

      const { listContactOpenExtensions } = await import('./outlook-graph-client.js');
      const r = await listContactOpenExtensions(token, 'c-1', undefined, { folderId: 'folder-1' });
      expect(r.ok).toBe(true);
      expect(urls[0]).toContain('/me/contactFolders/folder-1/contacts/c-1/extensions');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('listContactOpenExtensions uses childFolders segment when childFolderId set', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    const urls: string[] = [];
    const originalFetch = globalThis.fetch;

    try {
      globalThis.fetch = (async (input: string | URL | Request) => {
        urls.push(typeof input === 'string' ? input : input.toString());
        return new Response(JSON.stringify({ value: [] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }) as typeof fetch;

      const { listContactOpenExtensions } = await import('./outlook-graph-client.js');
      const r = await listContactOpenExtensions(token, 'c-2', undefined, {
        folderId: 'parent-f',
        childFolderId: 'child-f'
      });
      expect(r.ok).toBe(true);
      expect(urls[0]).toContain('/me/contactFolders/parent-f/childFolders/child-f/contacts/c-2/extensions');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});

describe('outlook mail folders and messages batch', () => {
  const originalFetch = globalThis.fetch;

  afterEach(() => {
    globalThis.fetch = originalFetch;
  });

  it('listChildMailFolders, create/update/delete mail folder', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
      const u = typeof input === 'string' ? input : input.toString();
      const m = (init?.method || 'GET').toUpperCase();
      if (m === 'DELETE') return new Response(null, { status: 204 });
      if (m === 'POST' || m === 'PATCH') {
        return new Response(JSON.stringify({ id: 'nf', displayName: 'N' }), {
          status: m === 'POST' ? 201 : 200,
          headers: { 'content-type': 'application/json' }
        });
      }
      return new Response(JSON.stringify({ value: [{ id: 'c1', displayName: 'Child' }] }), {
        status: 200,
        headers: { 'content-type': 'application/json' }
      });
    }) as typeof fetch;
    const o = await import('./outlook-graph-client.js');
    const ch = await o.listChildMailFolders(token, 'parent-f');
    expect(ch.ok).toBe(true);
    const cr = await o.createMailFolder(token, 'Sub', 'parent-f');
    expect(cr.ok).toBe(true);
    const up = await o.updateMailFolder(token, 'nf', { displayName: 'X' });
    expect(up.ok).toBe(true);
    const del = await o.deleteMailFolder(token, 'nf');
    expect(del.ok).toBe(true);
  });

  it('listMessagesInFolder, patchMailMessage, move/copy, sendMailMessage', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
      const u = typeof input === 'string' ? input : input.toString();
      const m = (init?.method || 'GET').toUpperCase();
      if (m === 'POST' && u.includes('/send')) {
        return new Response(null, { status: 202 });
      }
      if (m === 'POST') {
        return new Response(JSON.stringify({ id: 'm2' }), {
          status: 201,
          headers: { 'content-type': 'application/json' }
        });
      }
      if (m === 'PATCH') {
        return new Response(JSON.stringify({ id: 'm1', subject: 'P' }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }
      return new Response(JSON.stringify({ value: [{ id: 'm1', subject: 'A' }] }), {
        status: 200,
        headers: { 'content-type': 'application/json' }
      });
    }) as typeof fetch;
    const o = await import('./outlook-graph-client.js');
    const li = await o.listMessagesInFolder(token, 'inbox');
    expect(li.ok).toBe(true);
    const pa = await o.patchMailMessage(token, 'm1', { subject: 'P' });
    expect(pa.ok).toBe(true);
    const mv = await o.moveMailMessage(token, 'm1', 'dest-f');
    expect(mv.ok).toBe(true);
    const cp = await o.copyMailMessage(token, 'm1', 'dest-f');
    expect(cp.ok).toBe(true);
    const sm = await o.sendMailMessage(token, 'm1');
    expect(sm.ok).toBe(true);
  });

  it('mail message attachments list/get/download', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
      const u = typeof input === 'string' ? input : input.toString();
      const m = (init?.method || 'GET').toUpperCase();
      if (u.includes('/$value')) {
        return new Response(new Uint8Array([1, 2]), { status: 200 });
      }
      if (u.includes('/attachments/a1') && !u.includes('$value')) {
        return new Response(JSON.stringify({ id: 'a1', name: 'x' }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }
      return new Response(JSON.stringify({ value: [{ id: 'a1' }] }), {
        status: 200,
        headers: { 'content-type': 'application/json' }
      });
    }) as typeof fetch;
    const o = await import('./outlook-graph-client.js');
    const l = await o.listMailMessageAttachments(token, 'm1');
    expect(l.ok).toBe(true);
    const g = await o.getMailMessageAttachment(token, 'm1', 'a1');
    expect(g.ok).toBe(true);
    const d = await o.downloadMailMessageAttachmentBytes(token, 'm1', 'a1');
    expect(d.ok).toBe(true);
    expect(d.data?.length).toBe(2);
  });

  it('createMailReplyDraft, ReplyAll, Forward', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    globalThis.fetch = (async () =>
      new Response(JSON.stringify({ id: 'dr', isDraft: true }), {
        status: 201,
        headers: { 'content-type': 'application/json' }
      })) as typeof fetch;
    const o = await import('./outlook-graph-client.js');
    const r = await o.createMailReplyDraft(token, 'm1', undefined, 'c');
    expect(r.ok).toBe(true);
    const ra = await o.createMailReplyAllDraft(token, 'm1');
    expect(ra.ok).toBe(true);
    const fw = await o.createMailForwardDraft(token, 'm1', ['a@b.com']);
    expect(fw.ok).toBe(true);
  });
});

describe('outlook contacts batch', () => {
  const originalFetch = globalThis.fetch;

  afterEach(() => {
    globalThis.fetch = originalFetch;
  });

  it('contact folders CRUD and child folders', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
      const u = typeof input === 'string' ? input : input.toString();
      const m = (init?.method || 'GET').toUpperCase();
      if (m === 'DELETE') return new Response(null, { status: 204 });
      if (m === 'POST' || m === 'PATCH') {
        return new Response(JSON.stringify({ id: 'cf', displayName: 'F' }), {
          status: m === 'POST' ? 201 : 200,
          headers: { 'content-type': 'application/json' }
        });
      }
      return new Response(JSON.stringify({ value: [{ id: 'ch', displayName: 'H' }] }), {
        status: 200,
        headers: { 'content-type': 'application/json' }
      });
    }) as typeof fetch;
    const o = await import('./outlook-graph-client.js');
    const lf = await o.listContactFolders(token);
    expect(lf.ok).toBe(true);
    const gf = await o.getContactFolder(token, 'cf');
    expect(gf.ok).toBe(true);
    const cf = await o.createContactFolder(token, 'New');
    expect(cf.ok).toBe(true);
    const uf = await o.updateContactFolder(token, 'cf', { displayName: 'X' });
    expect(uf.ok).toBe(true);
    const ch = await o.listChildContactFolders(token, 'cf');
    expect(ch.ok).toBe(true);
    const df = await o.deleteContactFolder(token, 'cf');
    expect(df.ok).toBe(true);
  });

  it('contacts CRUD, folder list, delta, search', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
      const u = typeof input === 'string' ? input : input.toString();
      const m = (init?.method || 'GET').toUpperCase();
      if (m === 'DELETE') return new Response(null, { status: 204 });
      if (m === 'POST' || m === 'PATCH') {
        return new Response(JSON.stringify({ id: 'c1', displayName: 'Bob' }), {
          status: m === 'POST' ? 201 : 200,
          headers: { 'content-type': 'application/json' }
        });
      }
      if (u.includes('/delta')) {
        return new Response(JSON.stringify({ value: [] }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }
      return new Response(JSON.stringify({ value: [{ id: 'c1', displayName: 'Bob' }] }), {
        status: 200,
        headers: { 'content-type': 'application/json' }
      });
    }) as typeof fetch;
    const o = await import('./outlook-graph-client.js');
    const inf = await o.listContactsInFolder(token, 'folder-1');
    expect(inf.ok).toBe(true);
    const gc = await o.getContact(token, 'c1');
    expect(gc.ok).toBe(true);
    const cc = await o.createContact(token, { displayName: 'Bob', emailAddresses: [{ address: 'b@x.com' }] });
    expect(cc.ok).toBe(true);
    const uc = await o.updateContact(token, 'c1', { displayName: 'Bobby' });
    expect(uc.ok).toBe(true);
    const dp = await o.contactsDeltaPage(token, {});
    expect(dp.ok).toBe(true);
    const sr = await o.searchContacts(token, 'bob');
    expect(sr.ok).toBe(true);
    const dc = await o.deleteContact(token, 'c1');
    expect(dc.ok).toBe(true);
  });

  it('contact photo and attachments', async () => {
    process.env.GRAPH_BASE_URL = baseUrl;
    globalThis.fetch = (async (input: string | URL | Request, init?: RequestInit) => {
      const u = typeof input === 'string' ? input : input.toString();
      const m = (init?.method || 'GET').toUpperCase();
      if (m === 'DELETE') return new Response(null, { status: 204 });
      if (u.includes('/photo/$value') && m === 'GET') {
        return new Response(new Uint8Array([9]), { status: 200 });
      }
      if (u.includes('/photo/$value') && m === 'PUT') {
        return new Response(null, { status: 200 });
      }
      if (u.includes('/attachments/ca') && u.includes('/$value')) {
        return new Response(new Uint8Array([3]), { status: 200 });
      }
      if (u.includes('/attachments/ca') && m === 'GET') {
        return new Response(JSON.stringify({ id: 'ca', name: 'a' }), {
          status: 200,
          headers: { 'content-type': 'application/json' }
        });
      }
      if (m === 'POST') {
        return new Response(JSON.stringify({ id: 'ca' }), {
          status: 201,
          headers: { 'content-type': 'application/json' }
        });
      }
      return new Response(JSON.stringify({ value: [{ id: 'ca' }] }), {
        status: 200,
        headers: { 'content-type': 'application/json' }
      });
    }) as typeof fetch;
    const o = await import('./outlook-graph-client.js');
    const ph = await o.getContactPhotoBytes(token, 'c1');
    expect(ph.ok).toBe(true);
    const sp = await o.setContactPhoto(token, 'c1', new Uint8Array([1]), 'image/png');
    expect(sp.ok).toBe(true);
    const dp = await o.deleteContactPhoto(token, 'c1');
    expect(dp.ok).toBe(true);
    const la = await o.listContactAttachments(token, 'c1');
    expect(la.ok).toBe(true);
    const fa = await o.addFileAttachmentToContact(token, 'c1', {
      name: 'f',
      contentType: 'text/plain',
      contentBytes: 'YQ=='
    });
    expect(fa.ok).toBe(true);
    const ra = await o.addReferenceAttachmentToContact(token, 'c1', { name: 'r', sourceUrl: 'https://u' });
    expect(ra.ok).toBe(true);
    const ga = await o.getContactAttachment(token, 'c1', 'ca');
    expect(ga.ok).toBe(true);
    const dl = await o.downloadContactAttachmentBytes(token, 'c1', 'ca');
    expect(dl.ok).toBe(true);
    const da = await o.deleteContactAttachment(token, 'c1', 'ca');
    expect(da.ok).toBe(true);
  });
});
