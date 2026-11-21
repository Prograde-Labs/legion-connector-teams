// src/connector.test.ts
import { mkdtemp } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { TeamsConnector } from './connector.js';
import type { CollectiveLike, ConnectorRuntimeDeps } from '@legion-collective/core';
import type { TeamsAppLike } from './bot-pool.js';

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
