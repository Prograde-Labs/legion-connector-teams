// src/connector.test.ts
import { mkdtemp } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { describe, expect, it, vi } from 'vitest';
import { TeamsConnector } from './connector.js';
import type { CollectiveLike, ConnectorRuntimeDeps } from '@legion-collective/core';
import type { TeamsAppLike, TeamsAppFactory } from './bot-pool.js';
import { StateStore } from './state.js';

function setup() {
  const sent: Array<{ conversationId: string; activity: unknown }> = [];
  const app: TeamsAppLike = {
    async initialize() {},
    async start() {},
    async stop() {},
    async send(conversationId, activity) { sent.push({ conversationId, activity }); },
    api: { conversations: { create: async () => ({ id: 'conv-new' }) } },
    on() {},
  };
  const participants = [
    { id: 'assistant', name: 'Assistant', type: 'agent', tools: {}, status: 'active',
      identities: [{ connector: 'teams-bot', externalId: 'app' }] },
    { id: 'alice', name: 'Alice', type: 'user', tools: {}, status: 'active',
      identities: [{ connector: 'teams', externalId: 'aad-1' }] },
  ];
  const collective = {
    list: () => participants,
    get: (id: string) => participants.find((p) => p.id === id),
    findByIdentity: (connector: string, externalId: string) =>
      participants.find((p) =>
        p.identities?.some((i) => i.connector === connector && i.externalId === externalId)),
  } as unknown as CollectiveLike;
  const deps: ConnectorRuntimeDeps = {
    collective,
    secrets: {
      getSecret: async (id: string, key: string) =>
        id === 'assistant' && key === 'teams-bot/app'
          ? '{"appId":"app","clientSecret":"s","tenantId":"t"}'
          : null,
    },
  };
  return { app, sent, deps };
}

describe('TeamsConnector', () => {
  it('does not reconcile a paused discovery after stop begins', async () => {
    const { app, deps } = setup();
    const factory = vi.fn(() => app);
    const dir = await mkdtemp(join(tmpdir(), 'teams-rescan-'));
    const connector = new TeamsConnector({ stateDir: dir, port: 0 }, deps, factory);
    await connector.start({} as never);
    let release!: (value: string) => void;
    const secret = new Promise<string>(resolve => { release = resolve; });
    const discover = vi.spyOn(deps.secrets!, 'getSecret').mockReturnValue(secret);
    const scan = connector['rebuild']();
    const overlap = connector['rebuild']();
    const stopped = connector.stop();
    release('{"appId":"app","clientSecret":"s","tenantId":"t"}');
    await Promise.all([scan, overlap, stopped]);
    expect(discover).toHaveBeenCalledTimes(1);
    expect(factory).toHaveBeenCalledTimes(1);
  });

  it('only saves personal conversations with a usable id', async () => {
    const { app, deps } = setup();
    const factory = vi.fn<TeamsAppFactory>(() => app);
    const dir = await mkdtemp(join(tmpdir(), 'teams-install-'));
    const connector = new TeamsConnector({ stateDir: dir, transport: 'socket' }, deps, factory);
    await connector.start({} as never);
    const [entry, handlers] = factory.mock.calls[0];
    const context = (conversationType: string, id?: string) => ({
      activity: { from: { aadObjectId: 'aad-1' }, conversation: { conversationType, id } },
    } as never);
    try {
      const record = async () => (await new StateStore(dir).load()).conversations['app:alice'];
      await handlers.onInstall(entry, context('personal', 'personal-id'));
      const personal = await record();
      expect(personal.conversationId).toBe('personal-id');
      await handlers.onConversationUpdate(entry, context('groupChat', 'group-id'));
      expect(await record()).toEqual(personal);
      await handlers.onInstall(entry, context('personal'));
      expect(await record()).toEqual(personal);
    } finally { await connector.stop(); }
  });

  it('starts the pool and delivers outbound messages', async () => {
    const s = setup();
    const dir = await mkdtemp(join(tmpdir(), 'teams-conn-'));
    const c = new TeamsConnector(
      { stateDir: dir, port: 0, tenantId: 'tenant-1' },
      s.deps,
      () => s.app,
    );
    await c.start({ registry: { setActive() {}, clearActive() {} } } as never);
    await c.deliver({
      id: 'm1', conversationId: 'conv-1', senderId: 'assistant',
      recipientId: 'alice', content: 'hello', timestamp: 'now',
    });
    expect(s.sent).toHaveLength(1);
    expect(s.sent[0].conversationId).toBe('conv-new');
    await c.stop();
  });
});
