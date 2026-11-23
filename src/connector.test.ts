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
    { id: 'team-eng', name: 'Team Eng', type: 'user', tools: {}, status: 'active',
      identities: [{ connector: 'teams-group', externalId: '19:g@thread.v2' }] },
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
  return { app, sent, deps, participants };
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
    const context = (conversationType: string, id?: string, tenant = 't') => ({
      activity: { from: { aadObjectId: 'aad-1' }, conversation: { conversationType, id }, channelData: { tenant: { id: tenant } } },
    } as never);
    try {
      const record = async () => (await new StateStore(dir).load()).conversations['app:alice'];
      await handlers.onInstall(entry, context('personal', 'personal-id'));
      const personal = await record();
      expect(personal).toMatchObject({
        conversationId: 'personal-id',
        conversationType: 'personal',
        externalId: 'aad-1',
      });
      await handlers.onConversationUpdate(entry, context('groupChat', 'group-id'));
      expect(await record()).toEqual(personal);
      await handlers.onInstall(entry, context('personal'));
      expect(await record()).toEqual(personal);
      await handlers.onInstall(entry, context('personal', 'wrong-tenant', 'other'));
      expect(await record()).toEqual(personal);
    } finally { await connector.stop(); }
  });

  it('returns an Adaptive Card invoke response and rejects channels, retired or foreign actors', async () => {
    const { app, deps, participants } = setup();
    const factory = vi.fn<TeamsAppFactory>(() => app);
    const dir = await mkdtemp(join(tmpdir(), 'teams-approval-'));
    const connector = new TeamsConnector({ stateDir: dir, transport: 'socket' }, deps, factory);
    const callTool = vi.fn(async () => ({
      result: {
        status: 'success',
        data: {
          results: [{ approvalId: 'ap-1', outcome: 'approve' }],
          resumed: [{ conversationId: 'c1', response: 'resumed text' }],
        },
      },
    }));
    await connector.start({ callTool } as never);
    try {
      const [entry, handlers] = factory.mock.calls[0];
      const send = vi.fn();
      const invoke = (id: string, conversationType = 'personal') => handlers.onApproval(entry, { send, activity: {
        from: { aadObjectId: 'aad-1', name: 'Alice' }, channelData: { tenant: { id } },
        conversation: { conversationType },
        value: { action: { data: {
          approvalId: 'ap-1',
          tool: 'file_write',
          argsText: '{"path":"test.txt"}',
          requesterId: 'assistant',
        } } },
      } } as never, 'approve');
      expect(await invoke('other')).toHaveProperty('body.value', 'This bot is not available in your tenant.');
      expect(await invoke('t', 'channel')).toHaveProperty(
        'body.value', 'Only personal chats and group chats are supported.',
      );
      expect(callTool).not.toHaveBeenCalled();
      const approved = await invoke('t', 'groupChat');
      expect(approved).toMatchObject({ status: 200, body: {
        statusCode: 200, type: 'application/vnd.microsoft.card.adaptive',
      } });
      expect(JSON.stringify(approved.body)).toContain('Approved');
      expect(JSON.stringify(approved.body)).toContain('file_write');
      expect(JSON.stringify(approved.body)).toContain('Decided by');
      expect(JSON.stringify(approved.body)).toContain('Alice');
      expect((approved.body as { value?: { actions?: unknown[] } }).value?.actions ?? []).toHaveLength(0);
      participants[1].status = 'retired';
      expect(await invoke('t')).toHaveProperty('body.value', 'Not registered.');
      expect(callTool).toHaveBeenCalledOnce();
      expect(callTool).toHaveBeenCalledWith('alice', 'approval_response', {
        decisions: [{ approvalId: 'ap-1', decision: 'approve' }],
      });
      expect(send).toHaveBeenCalledOnce();
      expect(JSON.stringify(send.mock.calls[0][0])).toContain('resumed text');
    } finally { await connector.stop(); }
  });

  it('rejects group approval actions when group chats are disabled', async () => {
    const { app, deps } = setup();
    const factory = vi.fn<TeamsAppFactory>(() => app);
    const dir = await mkdtemp(join(tmpdir(), 'teams-approval-disabled-'));
    const connector = new TeamsConnector(
      { stateDir: dir, transport: 'socket', groupChats: false }, deps, factory,
    );
    const callTool = vi.fn(async () => ({ result: { status: 'success' }, conversationId: 'c' }));
    await connector.start({ callTool } as never);
    try {
      const [entry, handlers] = factory.mock.calls[0];
      const result = await handlers.onApproval(entry, {
        send: vi.fn(),
        activity: {
          from: { aadObjectId: 'aad-1' },
          channelData: { tenant: { id: 't' } },
          conversation: { conversationType: 'groupChat' },
          value: { action: { data: { approvalId: 'ap-1' } } },
        },
      } as never, 'approve');
      expect(result).toHaveProperty('body.value', 'Only personal chats are supported.');
      expect(callTool).not.toHaveBeenCalled();
    } finally { await connector.stop(); }
  });

  it('passes an optional rejection reason and shows it on the decided card', async () => {
    const { app, deps } = setup();
    const factory = vi.fn<TeamsAppFactory>(() => app);
    const dir = await mkdtemp(join(tmpdir(), 'teams-reject-'));
    const connector = new TeamsConnector({ stateDir: dir, transport: 'socket' }, deps, factory);
    const callTool = vi.fn(async () => ({
      result: {
        status: 'success',
        data: {
          results: [{ approvalId: 'ap-1', outcome: 'reject' }],
          resumed: [{ conversationId: 'c1', response: 'Understood.' }],
        },
      },
    }));
    await connector.start({ callTool } as never);
    try {
      const [entry, handlers] = factory.mock.calls[0];
      const send = vi.fn();
      const invoke = (reason?: string) => handlers.onApproval(entry, { send, activity: {
        from: { aadObjectId: 'aad-1' }, channelData: { tenant: { id: 't' } },
        conversation: { conversationType: 'personal' },
        value: { action: { data: {
          approvalId: 'ap-1',
          tool: 'file_write',
          argsText: '{"path":"test.txt"}',
          requesterId: 'assistant',
          ...(reason === undefined ? {} : { reason }),
        } } },
      } } as never, 'reject');
      const rejected = await invoke('  too risky  ');
      expect(callTool).toHaveBeenCalledWith('alice', 'approval_response', {
        decisions: [{ approvalId: 'ap-1', decision: 'reject', message: 'too risky' }],
      });
      expect(rejected).toMatchObject({ status: 200, body: {
        statusCode: 200, type: 'application/vnd.microsoft.card.adaptive',
      } });
      const body = JSON.stringify(rejected.body);
      expect(body).toContain('Rejected');
      expect(body).toContain('too risky');
      expect((rejected.body as { value?: { actions?: unknown[] } }).value?.actions ?? []).toHaveLength(0);
      expect(JSON.stringify(send.mock.calls[0][0])).toContain('Understood.');
    } finally { await connector.stop(); }
  });

  it('does not route outbound from a duplicate-app loser through the winner', async () => {
    const { app, deps, participants, sent } = setup();
    participants.push({ ...participants[0], id: 'duplicate' });
    const dir = await mkdtemp(join(tmpdir(), 'teams-duplicate-'));
    const connector = new TeamsConnector({ stateDir: dir, transport: 'socket' }, deps, () => app);
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    try {
      await connector.start({} as never);
      const message = { id: 'm1', conversationId: 'conv', recipientId: 'alice', content: 'hello', timestamp: 'now' };
      await connector.deliver({ ...message, senderId: 'duplicate' });
      expect(sent).toHaveLength(0);
      await connector.deliver({ ...message, senderId: 'assistant' });
      expect(sent).toHaveLength(1);
    } finally { await connector.stop(); warn.mockRestore(); }
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
    await expect(c.deliver({
      id: 'm1', conversationId: 'conv-1', senderId: 'assistant',
      recipientId: 'alice', content: 'hello', timestamp: 'now',
    })).resolves.toEqual({ status: 'delivered' });
    expect(s.sent).toHaveLength(1);
    expect(s.sent[0].conversationId).toBe('conv-new');
    await c.stop();
  });

  it('reports deliverability from Teams identities', () => {
    const { deps } = setup();
    const connector = new TeamsConnector({ stateDir: '/tmp/teams-can-deliver' }, deps);
    expect(connector.canDeliver('alice')).toBe(true);
    expect(connector.canDeliver('team-eng')).toBe(true);
    expect(connector.canDeliver('assistant')).toBe(false);
    expect(connector.canDeliver('nobody')).toBe(false);
  });
});
