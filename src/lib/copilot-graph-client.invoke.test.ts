/**
 * Exercises graphInvoke / graphInvokeText / callGraphAbsolute wrappers in copilot-graph-client.ts
 * with stubbed Graph layers so path construction and method selection are covered without network.
 */
import { afterAll, beforeAll, describe, expect, it, mock } from 'bun:test';
import type { GraphInvokeOptions } from './graph-advanced-client.js';

const graphAdvancedReal = await import('./graph-advanced-client.js');
const graphClientReal = await import('./graph-client.js');

function applyCopilotGraphStubs() {
  const graphInvoke = async (_token: string, opts: GraphInvokeOptions) => ({
    ok: true as const,
    data: { path: opts.path, method: opts.method }
  });
  const graphInvokeText = async (_token: string, opts: GraphInvokeOptions) => ({
    ok: true as const,
    data: `text:${opts.path}`
  });

  mock.module('./graph-advanced-client.js', () => ({
    ...graphAdvancedReal,
    graphInvoke,
    graphInvokeText
  }));

  mock.module('./graph-client.js', () => ({
    ...graphClientReal,
    callGraphAbsolute: async () => ({ ok: true as const, data: { via: 'absolute' } })
  }));
}

function restoreCopilotGraphModules() {
  mock.module('./graph-advanced-client.js', () => graphAdvancedReal);
  mock.module('./graph-client.js', () => graphClientReal);
}

describe('copilot-graph-client invoke wrappers', () => {
  let c: typeof import('./copilot-graph-client.js');

  beforeAll(async () => {
    applyCopilotGraphStubs();
    // Static import path so coverage attributes hits to copilot-graph-client.ts (query suffixes can split modules in tooling).
    c = await import('./copilot-graph-client.js');
  });

  afterAll(() => {
    restoreCopilotGraphModules();
  });

  const t = 'tok';
  const beta = true;
  const uid = 'user@contoso.com';
  const meet = 'meet-1';
  const insight = 'ins-1';
  const conv = 'conv-1';
  const msg = 'msg-1';
  const pkg = 'pkg-1';
  const agent = 'agent-1';
  const sub = 'sub-1';
  const aiUser = 'ai-user-1';
  const body: Record<string, unknown> = { x: 1 };

  it('covers graphInvoke-backed Copilot endpoints', async () => {
    await c.copilotRetrieval(t, body, beta);
    await c.copilotSearch(t, body, beta);
    await c.copilotSearchNextPage(t, 'https://graph.microsoft.com/beta/next');
    await c.copilotConversationCreate(t, beta);
    await c.copilotConversationChat(t, conv, body, beta);
    await c.copilotConversationChatOverStream(t, conv, body, beta);
    await c.copilotInteractionsExportList(t, uid, '$top=1', beta);
    await c.copilotInteractionsTenantExportList(t, '$top=1', beta);
    await c.copilotMeetingInsightsList(t, uid, meet, '$select=id', beta);
    await c.copilotMeetingInsightGet(t, uid, meet, insight, undefined, beta);
    await c.copilotMeetingAiInsightsCreate(t, uid, meet, body, beta);
    await c.copilotMeetingAiInsightsCount(t, uid, meet, '$top=1', beta);
    await c.copilotMeetingAiInsightPatch(t, uid, meet, insight, body, beta);
    await c.copilotMeetingAiInsightDelete(t, uid, meet, insight, beta);
    await c.copilotReportGet(t, 'getMicrosoft365CopilotUserCountSummary', 'D7', beta);
    await c.copilotPackagesList(t, '?$top=1');
    await c.copilotPackagesList(t, '$top=2');
    await c.copilotPackagesGet(t, pkg);
    await c.copilotPackagesUpdate(t, pkg, body);
    await c.copilotPackagesBlock(t, pkg);
    await c.copilotPackagesUnblock(t, pkg);
    await c.copilotPackagesReassign(t, pkg, uid);
    await c.copilotConversationsList(t, '$top=1', beta);
    await c.copilotConversationGet(t, conv, undefined, beta);
    await c.copilotConversationPatch(t, conv, body, beta);
    await c.copilotConversationDelete(t, conv, beta);
    await c.copilotConversationDeleteByThreadId(t, body, beta);
    await c.copilotConversationMessagesList(t, conv, undefined, beta);
    await c.copilotConversationMessageGet(t, conv, msg, undefined, beta);
    await c.copilotConversationMessageCreate(t, conv, body, beta);
    await c.copilotConversationMessagePatch(t, conv, msg, body, beta);
    await c.copilotConversationMessageDelete(t, conv, msg, beta);
    await c.copilotAgentsList(t, '$top=1', beta);
    await c.copilotAgentGet(t, agent, undefined, beta);
    await c.copilotSettingsGet(t, undefined, beta);
    await c.copilotSettingsPatch(t, body, beta);
    await c.copilotSettingsPeopleGet(t, undefined, beta);
    await c.copilotSettingsPeoplePatch(t, body, beta);
    await c.copilotSettingsEnhancedPersonalizationGet(t, undefined, beta);
    await c.copilotSettingsEnhancedPersonalizationPatch(t, body, beta);
    await c.copilotSettingsDelete(t, beta);
    await c.copilotSettingsPeopleDelete(t, beta);
    await c.copilotSettingsEnhancedPersonalizationDelete(t, beta);
    await c.copilotReportsNavGet(t, undefined, beta);
    await c.copilotReportsNavPatch(t, body, beta);
    await c.copilotReportsNavDelete(t, beta);
    await c.copilotAdminSettingsGet(t, undefined);
    await c.copilotAdminSettingsPatch(t, body);
    await c.copilotAdminLimitedModeGet(t, undefined);
    await c.copilotAdminLimitedModePatch(t, body);
    await c.copilotPackagesCreate(t, body);
    await c.copilotPackagesDelete(t, pkg);
    await c.copilotPackageZipDelete(t, pkg);
    await c.copilotAdminSettingsDelete(t);
    await c.copilotAdminLimitedModeDelete(t);
    await c.copilotRealtimeActivityFeedGet(t, undefined, beta);
    await c.copilotRealtimeMeetingsList(t, undefined, beta);
    await c.copilotRealtimeMeetingCreate(t, body, beta);
    await c.copilotRealtimeMeetingGet(t, meet, undefined, beta);
    await c.copilotRealtimeMeetingPatch(t, meet, body, beta);
    await c.copilotRealtimeMeetingDelete(t, meet, beta);
    await c.copilotRealtimeSubscriptionsList(t, undefined, beta);
    await c.copilotRealtimeSubscriptionCreate(t, body, beta);
    await c.copilotRealtimeSubscriptionGet(t, sub, undefined, beta);
    await c.copilotRealtimeSubscriptionPatch(t, sub, body, beta);
    await c.copilotRealtimeSubscriptionDelete(t, sub, beta);
    await c.copilotRealtimeSubscriptionGetArtifacts(t, sub, body, beta);
    await c.copilotRealtimeTranscriptsList(t, meet, undefined, beta);
    await c.copilotRealtimeTranscriptCreate(t, meet, body, beta);
    await c.copilotRealtimeTranscriptGet(t, meet, 'tr-1', undefined, beta);
    await c.copilotRealtimeTranscriptPatch(t, meet, 'tr-1', body, beta);
    await c.copilotRealtimeTranscriptDelete(t, meet, 'tr-1', beta);
    await c.copilotRootGet(t, undefined, beta);
    await c.copilotRootPatch(t, body, beta);
    await c.copilotAdminNavGet(t, undefined);
    await c.copilotAdminNavPatch(t, body);
    await c.copilotAdminNavDelete(t);
    await c.copilotAdminCatalogGet(t, undefined);
    await c.copilotAdminCatalogPatch(t, body);
    await c.copilotAdminCatalogDelete(t);
    await c.copilotPackagesCount(t, undefined);
    await c.copilotCommunicationsGet(t, undefined, beta);
    await c.copilotCommunicationsPatch(t, body, beta);
    await c.copilotCommunicationsDelete(t, beta);
    await c.copilotRealtimeActivityFeedPatch(t, body, beta);
    await c.copilotRealtimeActivityFeedDelete(t, beta);
    await c.copilotConversationsCount(t, undefined, beta);
    await c.copilotConversationMessagesCount(t, conv, undefined, beta);
    await c.copilotAgentsCount(t, undefined, beta);
    await c.copilotRealtimeMeetingsCount(t, undefined, beta);
    await c.copilotRealtimeTranscriptsCount(t, meet, undefined, beta);
    await c.copilotRealtimeSubscriptionsCount(t, undefined, beta);
    await c.copilotInteractionHistoryNavGet(t, undefined, beta);
    await c.copilotInteractionHistoryNavPatch(t, body, beta);
    await c.copilotInteractionHistoryNavDelete(t, beta);
    await c.copilotAiUsersList(t, undefined, beta);
    await c.copilotAiUsersCount(t, undefined, beta);
    await c.copilotAiUserCreate(t, body, beta);
    await c.copilotAiUserGet(t, aiUser, undefined, beta);
    await c.copilotAiUserPatch(t, aiUser, body, beta);
    await c.copilotAiUserDelete(t, aiUser, beta);
    await c.copilotAiUserInteractionHistoryGet(t, aiUser, undefined, beta);
    await c.copilotAiUserInteractionHistoryPatch(t, aiUser, body, beta);
    await c.copilotAiUserInteractionHistoryDelete(t, aiUser, beta);
    await c.copilotAiUserOnlineMeetingsList(t, aiUser, undefined, beta);
    await c.copilotAiUserOnlineMeetingsCount(t, aiUser, undefined, beta);
    await c.copilotAiUserOnlineMeetingCreate(t, aiUser, body, beta);
    await c.copilotAiUserOnlineMeetingGet(t, aiUser, meet, undefined, beta);
    await c.copilotAiUserOnlineMeetingPatch(t, aiUser, meet, body, beta);
    await c.copilotAiUserOnlineMeetingDelete(t, aiUser, meet, beta);
  });

  it('copilotSearchNextPage maps GraphApiError to error response', async () => {
    mock.module('./graph-client.js', () => ({
      ...graphClientReal,
      callGraphAbsolute: async () => {
        throw new graphClientReal.GraphApiError('nope', 'X', 403);
      }
    }));
    const mod = await import(`./copilot-graph-client.js?copilotErr=${Date.now()}`);
    const r = await mod.copilotSearchNextPage(t, 'https://graph.microsoft.com/beta/page2');
    expect(r.ok).toBe(false);
    expect(r.error?.message).toContain('nope');
    mock.module('./graph-client.js', () => ({
      ...graphClientReal,
      callGraphAbsolute: async () => ({ ok: true as const, data: {} })
    }));
  });

  it('copilotSearchNextPage maps generic errors', async () => {
    mock.module('./graph-client.js', () => ({
      ...graphClientReal,
      callGraphAbsolute: async () => {
        throw new Error('boom');
      }
    }));
    const mod = await import(`./copilot-graph-client.js?copilotErr2=${Date.now()}`);
    const r = await mod.copilotSearchNextPage(t, 'https://graph.microsoft.com/beta/page2');
    expect(r.ok).toBe(false);
    expect(r.error?.message).toContain('boom');
    mock.module('./graph-client.js', () => ({
      ...graphClientReal,
      callGraphAbsolute: async () => ({ ok: true as const, data: {} })
    }));
  });
});

describe('copilot package zip fetch paths', () => {
  const originalFetch = globalThis.fetch;

  afterAll(() => {
    globalThis.fetch = originalFetch;
    restoreCopilotGraphModules();
  });

  it('copilotPackageZipDownload returns bytes on 200', async () => {
    applyCopilotGraphStubs();
    await import(`./copilot-graph-client.js?zipDl=${Date.now()}`);
    globalThis.fetch = (async () =>
      new Response(new Uint8Array([1, 2, 3]), {
        status: 200,
        headers: { 'content-type': 'application/octet-stream' }
      })) as unknown as typeof fetch;

    const mod = await import(`./copilot-graph-client.js?zipDl2=${Date.now()}`);
    const r = await mod.copilotPackageZipDownload('tok', 'pkg-1');
    expect(r.ok).toBe(true);
    expect(r.data?.length).toBe(3);
  });

  it('copilotPackageZipDownload returns graphError on failed JSON body', async () => {
    applyCopilotGraphStubs();
    globalThis.fetch = (async () =>
      new Response(JSON.stringify({ error: { message: 'bad' } }), {
        status: 400,
        headers: { 'content-type': 'application/json' }
      })) as unknown as typeof fetch;

    const mod = await import(`./copilot-graph-client.js?zipDl3=${Date.now()}`);
    const r = await mod.copilotPackageZipDownload('tok', 'pkg-1');
    expect(r.ok).toBe(false);
    expect(r.error?.message).toBe('bad');
  });

  it('copilotPackageZipUpload handles 204 and JSON success', async () => {
    applyCopilotGraphStubs();
    let n = 0;
    globalThis.fetch = (async () => {
      n += 1;
      if (n === 1) {
        return new Response(null, { status: 204 });
      }
      return new Response(JSON.stringify({ ok: true }), {
        status: 200,
        headers: { 'content-type': 'application/json' }
      });
    }) as unknown as typeof fetch;

    const mod = await import(`./copilot-graph-client.js?zipUl=${Date.now()}`);
    const a = await mod.copilotPackageZipUpload('tok', 'p', new Uint8Array([9]), 'application/octet-stream');
    expect(a.ok).toBe(true);
    const b = await mod.copilotPackageZipUpload('tok', 'p', new Uint8Array([9]), 'application/octet-stream');
    expect(b.ok).toBe(true);
    expect(b.data).toEqual({ ok: true });
  });
});
