import { describe, expect, it } from 'bun:test';
import { addDelegate, getDelegates, removeDelegate, updateDelegate } from './delegate-client.js';

describe('delegate-client', () => {
  const token = 'test-token';

  it('getDelegates parses SOAP response properly', async () => {
    const fetchCalls: any[] = [];
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async (input, init) => {
        fetchCalls.push({ input, init });
        const xml = `
          <s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/">
            <s:Body>
              <m:GetDelegateResponse xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
                <m:ResponseMessages>
                  <m:DelegateUserResponseMessageType ResponseClass="Success">
                    <m:DelegateUser>
                      <t:UserId>
                        <t:PrimarySmtpAddress>del@example.com</t:PrimarySmtpAddress>
                      </t:UserId>
                      <t:DelegatePermissions>
                        <t:CalendarFolderPermissionLevel>Editor</t:CalendarFolderPermissionLevel>
                      </t:DelegatePermissions>
                      <t:ViewPrivateItems>true</t:ViewPrivateItems>
                    </m:DelegateUser>
                  </m:DelegateUserResponseMessageType>
                </m:ResponseMessages>
                <m:DeliverMeetingRequests>DelegatesAndSendInformationToMe</m:DeliverMeetingRequests>
              </m:GetDelegateResponse>
            </s:Body>
          </s:Envelope>
        `;
        return new Response(xml, { status: 200, headers: { 'content-type': 'text/xml' } });
      }) as typeof fetch;

      const res = await getDelegates(token, 'me@example.com');
      expect(res.ok).toBe(true);
      expect(res.data?.length).toBe(1);
      expect(res.data?.[0].userId).toBe('del@example.com');
      expect(res.data?.[0].permissions.calendar).toBe('Editor');
      expect(res.data?.[0].viewPrivateItems).toBe(true);
      expect(res.data?.[0].deliverMeetingRequests).toBe('DelegatesAndSendInformationToMe');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('addDelegate generates correct SOAP body', async () => {
    const fetchCalls: any[] = [];
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async (input, init) => {
        fetchCalls.push({ input, init });
        const xml = `<m:AddDelegateResponse ResponseClass="Success"></m:AddDelegateResponse>`;
        return new Response(xml, { status: 200, headers: { 'content-type': 'text/xml' } });
      }) as typeof fetch;

      await addDelegate({
        token,
        delegateEmail: 'del@example.com',
        permissions: { inbox: 'Reviewer' },
        deliverMeetingRequests: 'DelegatesAndSendInformationToMe'
      });

      const body = fetchCalls[0].init.body as string;
      expect(body).toContain('<m:DeliverMeetingRequests>DelegatesAndSendInformationToMe</m:DeliverMeetingRequests>');
      expect(body).toContain('<t:InboxFolderPermissionLevel>Reviewer</t:InboxFolderPermissionLevel>');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('updateDelegate returns NO_UPDATES when nothing to change', async () => {
    const r = await updateDelegate({ token, delegateEmail: 'd@x.com' });
    expect(r.ok).toBe(false);
    expect(r.error?.code).toBe('NO_UPDATES');
  });

  it('updateDelegate and removeDelegate call EWS', async () => {
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async () => {
        const xml = `
          <m:UpdateDelegateResponse xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
            <m:ResponseMessages>
              <m:DelegateUserResponseMessageType ResponseClass="Success">
                <m:DelegateUser>
                  <t:UserId xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
                    <t:PrimarySmtpAddress>del@example.com</t:PrimarySmtpAddress>
                  </t:UserId>
                  <t:DelegatePermissions xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
                    <t:TasksFolderPermissionLevel>Editor</t:TasksFolderPermissionLevel>
                  </t:DelegatePermissions>
                </m:DelegateUser>
              </m:DelegateUserResponseMessageType>
            </m:ResponseMessages>
          </m:UpdateDelegateResponse>`;
        return new Response(xml, { status: 200, headers: { 'content-type': 'text/xml' } });
      }) as typeof fetch;

      const u = await updateDelegate({
        token,
        delegateEmail: 'del@example.com',
        permissions: { tasks: 'Editor' },
        viewPrivateItems: true
      });
      expect(u.ok).toBe(true);
      expect(u.data?.permissions.tasks).toBe('Editor');
    } finally {
      globalThis.fetch = originalFetch;
    }

    try {
      globalThis.fetch = (async () =>
        new Response(
          `<m:RemoveDelegateResponse ResponseClass="Success" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"/>`,
          { status: 200, headers: { 'content-type': 'text/xml' } }
        )) as typeof fetch;
      const rm = await removeDelegate({ token, delegateEmail: 'del@example.com' });
      expect(rm.ok).toBe(true);
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('getDelegates uses SmtpAddress when PrimarySmtpAddress missing', async () => {
    const originalFetch = globalThis.fetch;
    try {
      globalThis.fetch = (async () => {
        const xml = `
          <s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/">
            <s:Body>
              <m:GetDelegateResponse xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
                <m:ResponseMessages>
                  <m:DelegateUserResponseMessageType ResponseClass="Success">
                    <m:DelegateUser>
                      <t:UserId>
                        <t:SmtpAddress>alt@example.com</t:SmtpAddress>
                      </t:UserId>
                      <t:DelegatePermissions />
                    </m:DelegateUser>
                  </m:DelegateUserResponseMessageType>
                </m:ResponseMessages>
              </m:GetDelegateResponse>
            </s:Body>
          </s:Envelope>`;
        return new Response(xml, { status: 200, headers: { 'content-type': 'text/xml' } });
      }) as typeof fetch;
      const res = await getDelegates(token);
      expect(res.ok).toBe(true);
      expect(res.data?.[0].userId).toBe('alt@example.com');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});
