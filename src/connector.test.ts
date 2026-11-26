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
  const updates: Array<{ conversationId: string; activityId: string; activity: unknown }> = [];
  const app: TeamsAppLike = {
    async initialize() {},
    async start() {},
    async stop() {},
    async send(conversationId, activity) {
      sent.push({ conversationId, activity });
      return { id: 'activity-1' };
    },
    api: {
      conversations: {
        create: async () => ({ id: 'conv-new' }),
        updateActivity: async (conversationId, activityId, activity) => {
          updates.push({ conversationId, activityId, activity });
        },
      },
      reactions: { add: async () => undefined, delete: async () => undefined },
    },
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
  return { app, sent, updates, deps, participants };
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
  it('keeps the configured fallback reachable without a Teams identity', () => {
    const { deps, participants } = setup();
    participants.push({ ...participants[1], id: 'fallback-user', identities: [] });
    const connector = new TeamsConnector({ stateDir: '/tmp/teams-can-deliver', fallbackUserId: 'fallback-user' }, deps);
    expect(connector.canDeliver('fallback-user')).toBe(true);
    expect(connector.canDeliver('assistant')).toBe(false);
  });

  describe('card tool actions', () => {
    it('runs the tool as the clicker and updates the card in place', async () => {
      const { app, deps, updates } = setup();
      const factory = vi.fn<TeamsAppFactory>(() => app);
      const dir = await mkdtemp(join(tmpdir(), 'teams-tool-action-'));
      const connector = new TeamsConnector({ stateDir: dir, transport: 'socket' }, deps, factory);
      const callTool = vi.fn(async () => ({
        result: { status: 'success', data: { reaction: 'like' } },
        conversationId: 'c1',
      }));
      await connector.start({ callTool } as never);
      try {
        const [entry, handlers] = factory.mock.calls[0];
        const response = await handlers.onToolAction(entry, {
          send: vi.fn(),
          activity: {
            from: { aadObjectId: 'aad-1', name: 'Alice' },
            channelData: { tenant: { id: 't' } },
            conversation: { conversationType: 'personal', id: 'conv-1' },
            replyToId: 'card-1',
            value: {
              action: {
                data: {
                  action: 'tool',
                  tool: 'teams__post_reaction',
                  args: { reaction: 'like' },
                },
              },
            },
          },
        } as never);
        expect(response).toMatchObject({
          status: 200,
          body: { statusCode: 200, type: 'application/vnd.microsoft.card.adaptive' },
        });
        expect(JSON.stringify(response.body)).toContain('Running teams__post_reaction');
        expect(callTool).toHaveBeenCalledWith(
          'alice',
          'teams__post_reaction',
          { reaction: 'like' },
          { conversationId: '' },
        );
        await vi.waitFor(() => expect(updates).toHaveLength(1));
        expect(updates[0]).toMatchObject({ conversationId: 'conv-1', activityId: 'card-1' });
        expect(JSON.stringify(updates[0].activity)).toContain('like');
      } finally {
        await connector.stop();
      }
    });

    it('rejects an unknown clicker without calling the tool', async () => {
      const { app, deps } = setup();
      const factory = vi.fn<TeamsAppFactory>(() => app);
      const dir = await mkdtemp(join(tmpdir(), 'teams-tool-unknown-'));
      const connector = new TeamsConnector({ stateDir: dir, transport: 'socket' }, deps, factory);
      const callTool = vi.fn(async () => ({ result: { status: 'success' }, conversationId: 'c' }));
      await connector.start({ callTool } as never);
      try {
        const [entry, handlers] = factory.mock.calls[0];
        const response = await handlers.onToolAction(entry, {
          send: vi.fn(),
          activity: {
            from: { aadObjectId: 'aad-2' },
            channelData: { tenant: { id: 't' } },
            conversation: { conversationType: 'personal', id: 'conv-1' },
            value: { action: { data: { action: 'tool', tool: 'teams__post_reaction' } } },
          },
        } as never);
        expect(response).toHaveProperty('body.value', 'Not registered.');
        expect(callTool).not.toHaveBeenCalled();
      } finally {
        await connector.stop();
      }
    });

    it('rejects a missing tool and malformed args', async () => {
      const { app, deps } = setup();
      const factory = vi.fn<TeamsAppFactory>(() => app);
      const dir = await mkdtemp(join(tmpdir(), 'teams-tool-invalid-'));
      const connector = new TeamsConnector({ stateDir: dir, transport: 'socket' }, deps, factory);
      const callTool = vi.fn(async () => ({ result: { status: 'success' }, conversationId: 'c' }));
      await connector.start({ callTool } as never);
      try {
        const [entry, handlers] = factory.mock.calls[0];
        const invoke = (data: unknown) => handlers.onToolAction(entry, {
          send: vi.fn(),
          activity: {
            from: { aadObjectId: 'aad-1' },
            channelData: { tenant: { id: 't' } },
            conversation: { conversationType: 'personal', id: 'conv-1' },
            value: { action: { data } },
          },
        } as never);
        expect(await invoke({ action: 'tool' })).toHaveProperty('body.value', 'Missing tool.');
        expect(await invoke({ action: 'tool', tool: 't', args: [] }))
          .toHaveProperty('body.value', 'Invalid args.');
        expect(callTool).not.toHaveBeenCalled();
      } finally {
        await connector.stop();
      }
    });

    it('rejects an over-long tool name without calling the tool', async () => {
      const { app, deps } = setup();
      const factory = vi.fn<TeamsAppFactory>(() => app);
      const dir = await mkdtemp(join(tmpdir(), 'teams-tool-long-'));
      const connector = new TeamsConnector({ stateDir: dir, transport: 'socket' }, deps, factory);
      const callTool = vi.fn(async () => ({ result: { status: 'success' }, conversationId: 'c' }));
      await connector.start({ callTool } as never);
      try {
        const [entry, handlers] = factory.mock.calls[0];
        const response = await handlers.onToolAction(entry, {
          send: vi.fn(),
          activity: {
            from: { aadObjectId: 'aad-1' },
            channelData: { tenant: { id: 't' } },
            conversation: { conversationType: 'personal', id: 'conv-1' },
            value: { action: { data: { action: 'tool', tool: 't'.repeat(129) } } },
          },
        } as never);
        expect(response).toHaveProperty('body.value', 'Invalid tool.');
        expect(callTool).not.toHaveBeenCalled();
      } finally {
        await connector.stop();
      }
    });

    it('rejects tenant and conversation-type mismatches', async () => {
      const { app, deps } = setup();
      const factory = vi.fn<TeamsAppFactory>(() => app);
      const dir = await mkdtemp(join(tmpdir(), 'teams-tool-scope-'));
      const connector = new TeamsConnector({ stateDir: dir, transport: 'socket' }, deps, factory);
      const callTool = vi.fn(async () => ({ result: { status: 'success' }, conversationId: 'c' }));
      await connector.start({ callTool } as never);
      try {
        const [entry, handlers] = factory.mock.calls[0];
        const invoke = (tenant: string, conversationType: string) => handlers.onToolAction(entry, {
          send: vi.fn(),
          activity: {
            from: { aadObjectId: 'aad-1' },
            channelData: { tenant: { id: tenant } },
            conversation: { conversationType, id: 'conv-1' },
            value: { action: { data: { action: 'tool', tool: 'teams__post_reaction' } } },
          },
        } as never);
        expect(await invoke('other', 'personal'))
          .toHaveProperty('body.value', 'This bot is not available in your tenant.');
        expect(await invoke('t', 'channel'))
          .toHaveProperty('body.value', 'Only personal chats and group chats are supported.');
        expect(callTool).not.toHaveBeenCalled();
      } finally {
        await connector.stop();
      }
    });

    it('renders requires_approval as a result card', async () => {
      const { app, deps, updates } = setup();
      const factory = vi.fn<TeamsAppFactory>(() => app);
      const dir = await mkdtemp(join(tmpdir(), 'teams-tool-approval-'));
      const connector = new TeamsConnector({ stateDir: dir, transport: 'socket' }, deps, factory);
      const callTool = vi.fn(async () => ({
        result: { status: 'error', error: 'requires_approval' },
        conversationId: 'c1',
      }));
      await connector.start({ callTool } as never);
      try {
        const [entry, handlers] = factory.mock.calls[0];
        await handlers.onToolAction(entry, {
          send: vi.fn(),
          activity: {
            from: { aadObjectId: 'aad-1', name: 'Alice' },
            channelData: { tenant: { id: 't' } },
            conversation: { conversationType: 'personal', id: 'conv-1' },
            replyToId: 'card-1',
            value: { action: { data: { action: 'tool', tool: 'teams__post_reaction' } } },
          },
        } as never);
        await vi.waitFor(() => expect(updates).toHaveLength(1));
        expect(JSON.stringify(updates[0].activity)).toContain('requires approval');
      } finally {
        await connector.stop();
      }
    });

    it('sends a follow-up when replyToId is absent', async () => {
      const { app, deps, sent } = setup();
      const factory = vi.fn<TeamsAppFactory>(() => app);
      const dir = await mkdtemp(join(tmpdir(), 'teams-tool-followup-'));
      const connector = new TeamsConnector({ stateDir: dir, transport: 'socket' }, deps, factory);
      const callTool = vi.fn(async () => ({
        result: { status: 'success', data: { reaction: 'like' } },
        conversationId: 'c1',
      }));
      await connector.start({ callTool } as never);
      try {
        const [entry, handlers] = factory.mock.calls[0];
        await handlers.onToolAction(entry, {
          send: vi.fn(),
          activity: {
            from: { aadObjectId: 'aad-1' },
            channelData: { tenant: { id: 't' } },
            conversation: { conversationType: 'personal', id: 'conv-1' },
            value: {
              action: {
                data: { action: 'tool', tool: 'teams__post_reaction', args: { reaction: 'like' } },
              },
            },
          },
        } as never);
        await vi.waitFor(() => expect(sent).toHaveLength(1));
        expect(sent[0].conversationId).toBe('conv-1');
        expect(JSON.stringify(sent[0].activity)).toContain('like');
      } finally {
        await connector.stop();
      }
    });

    it('reuses the mapped Legion conversation', async () => {
      const { app, deps, updates } = setup();
      const factory = vi.fn<TeamsAppFactory>(() => app);
      const dir = await mkdtemp(join(tmpdir(), 'teams-tool-mapped-'));
      const connector = new TeamsConnector({ stateDir: dir, transport: 'socket' }, deps, factory);
      const callTool = vi.fn(async () => ({
        result: { status: 'success' },
        conversationId: 'legion-1',
      }));
      await connector.start({ callTool } as never);
      try {
        connector['state']!.legionConversations['app:alice'] = 'legion-1';
        const [entry, handlers] = factory.mock.calls[0];
        await handlers.onToolAction(entry, {
          send: vi.fn(),
          activity: {
            from: { aadObjectId: 'aad-1' },
            channelData: { tenant: { id: 't' } },
            conversation: { conversationType: 'personal', id: 'conv-1' },
            replyToId: 'card-1',
            value: { action: { data: { action: 'tool', tool: 'teams__post_reaction' } } },
          },
        } as never);
        expect(callTool).toHaveBeenCalledWith(
          'alice',
          'teams__post_reaction',
          {},
          { conversationId: 'legion-1' },
        );
        await vi.waitFor(() => expect(updates).toHaveLength(1));
      } finally {
        await connector.stop();
      }
    });

    it('renders a tool exception as an error card', async () => {
      const { app, deps, updates } = setup();
      const factory = vi.fn<TeamsAppFactory>(() => app);
      const dir = await mkdtemp(join(tmpdir(), 'teams-tool-error-'));
      const connector = new TeamsConnector({ stateDir: dir, transport: 'socket' }, deps, factory);
      const callTool = vi.fn(async () => {
        throw new Error('boom');
      });
      await connector.start({ callTool } as never);
      try {
        const [entry, handlers] = factory.mock.calls[0];
        await handlers.onToolAction(entry, {
          send: vi.fn(),
          activity: {
            from: { aadObjectId: 'aad-1' },
            channelData: { tenant: { id: 't' } },
            conversation: { conversationType: 'personal', id: 'conv-1' },
            replyToId: 'card-1',
            value: { action: { data: { action: 'tool', tool: 'teams__post_reaction' } } },
          },
        } as never);
        await vi.waitFor(() => expect(updates).toHaveLength(1));
        expect(JSON.stringify(updates[0].activity)).toContain('boom');
      } finally {
        await connector.stop();
      }
    });

    it('runs a group-chat tool as the clicker in the mapped group conversation', async () => {
      const { app, deps, updates } = setup();
      const factory = vi.fn<TeamsAppFactory>(() => app);
      const dir = await mkdtemp(join(tmpdir(), 'teams-tool-group-'));
      const connector = new TeamsConnector({ stateDir: dir, transport: 'socket' }, deps, factory);
      const callTool = vi.fn(async () => ({
        result: { status: 'success', data: { reaction: 'like' } },
        conversationId: 'legion-group',
      }));
      await connector.start({ callTool } as never);
      try {
        connector['state']!.legionConversations['app:team-eng:19:g@thread.v2'] = 'legion-group';
        const [entry, handlers] = factory.mock.calls[0];
        await handlers.onToolAction(entry, {
          send: vi.fn(),
          activity: {
            from: { aadObjectId: 'aad-1', name: 'Alice' },
            channelData: { tenant: { id: 't' } },
            conversation: { conversationType: 'groupChat', id: '19:g@thread.v2' },
            replyToId: 'card-1',
            value: {
              action: {
                data: { action: 'tool', tool: 'teams__post_reaction', args: { reaction: 'like' } },
              },
            },
          },
        } as never);
        expect(callTool).toHaveBeenCalledWith(
          'alice',
          'teams__post_reaction',
          { reaction: 'like' },
          { conversationId: 'legion-group' },
        );
        await vi.waitFor(() => expect(updates).toHaveLength(1));
        expect(updates[0]).toMatchObject({ conversationId: '19:g@thread.v2', activityId: 'card-1' });
      } finally {
        await connector.stop();
      }
    });

    it('falls back to an empty conversation id when no group participant resolves', async () => {
      const { app, deps, updates } = setup();
      const factory = vi.fn<TeamsAppFactory>(() => app);
      const dir = await mkdtemp(join(tmpdir(), 'teams-tool-group-unknown-'));
      const connector = new TeamsConnector({ stateDir: dir, transport: 'socket' }, deps, factory);
      const callTool = vi.fn(async () => ({ result: { status: 'success' }, conversationId: 'c' }));
      await connector.start({ callTool } as never);
      try {
        const [entry, handlers] = factory.mock.calls[0];
        await handlers.onToolAction(entry, {
          send: vi.fn(),
          activity: {
            from: { aadObjectId: 'aad-1' },
            channelData: { tenant: { id: 't' } },
            conversation: { conversationType: 'groupChat', id: '19:unknown@thread.v2' },
            replyToId: 'card-1',
            value: { action: { data: { action: 'tool', tool: 'teams__post_reaction' } } },
          },
        } as never);
        expect(callTool).toHaveBeenCalledWith(
          'alice',
          'teams__post_reaction',
          {},
          { conversationId: '' },
        );
        await vi.waitFor(() => expect(updates).toHaveLength(1));
      } finally {
        await connector.stop();
      }
    });

    it('logs and swallows a result-card build failure instead of rejecting', async () => {
      const { app, deps, updates } = setup();
      const factory = vi.fn<TeamsAppFactory>(() => app);
      const dir = await mkdtemp(join(tmpdir(), 'teams-tool-circular-'));
      const connector = new TeamsConnector({ stateDir: dir, transport: 'socket' }, deps, factory);
      const circular: Record<string, unknown> = {};
      circular.self = circular;
      const callTool = vi.fn(async () => ({
        result: { status: 'success', data: [circular] },
        conversationId: 'c1',
      }));
      await connector.start({ callTool } as never);
      const errorSpy = vi.spyOn(console, 'error').mockImplementation(() => {});
      try {
        const [entry, handlers] = factory.mock.calls[0];
        const response = await handlers.onToolAction(entry, {
          send: vi.fn(),
          activity: {
            from: { aadObjectId: 'aad-1' },
            channelData: { tenant: { id: 't' } },
            conversation: { conversationType: 'personal', id: 'conv-1' },
            replyToId: 'card-1',
            value: { action: { data: { action: 'tool', tool: 'teams__post_reaction' } } },
          },
        } as never);
        expect(response).toMatchObject({
          status: 200,
          body: { statusCode: 200, type: 'application/vnd.microsoft.card.adaptive' },
        });
        await vi.waitFor(() =>
          expect(errorSpy).toHaveBeenCalledWith(
            expect.stringContaining('Teams card action result'),
            expect.stringContaining('circular'),
          ),
        );
        expect(updates).toHaveLength(0);
      } finally {
        await connector.stop();
        errorSpy.mockRestore();
      }
    });
  });

  it('switches the chat to a chosen conversation and updates the card', async () => {
    const { app, deps, updates } = setup();
    const factory = vi.fn<TeamsAppFactory>(() => app);
    const dir = await mkdtemp(join(tmpdir(), 'teams-switch-'));
    const connector = new TeamsConnector({ stateDir: dir, transport: 'socket' }, deps, factory);
    const callTool = vi.fn(async () => ({
      result: { status: 'success', data: { conversations: [
        { id: 'legion-1', resolvedTitle: 'First', messageCount: 4,
          updatedAt: '2026-11-24T00:00:00.000Z', status: 'active' },
        { id: 'legion-2', resolvedTitle: 'Second', messageCount: 2,
          updatedAt: '2026-11-25T00:00:00.000Z', status: 'active' },
      ] } },
      conversationId: 'c',
    }));
    await connector.start({ callTool } as never);
    try {
      connector['state']!.legionConversations['app:alice'] = 'legion-1';
      connector['state']!.legionConversationTargets['legion-1'] = 'app:alice';
      const [entry, handlers] = factory.mock.calls[0];
      const response = await handlers.onConversationAction(entry, {
        send: vi.fn(),
        activity: {
          from: { aadObjectId: 'aad-1' },
          channelData: { tenant: { id: 't' } },
          conversation: { conversationType: 'personal', id: 'conv-1' },
          replyToId: 'card-1',
          value: { action: { data: {
            action: 'switch', conversationId: 'legion-2', key: 'app:alice',
          } } },
        },
      } as never);
      expect(response).toMatchObject({
        status: 200,
        body: { statusCode: 200, type: 'application/vnd.microsoft.card.adaptive' },
      });
      expect(JSON.stringify(response.body)).toContain('Running switch');
      await vi.waitFor(() => expect(updates).toHaveLength(1));
      expect(updates[0]).toMatchObject({ conversationId: 'conv-1', activityId: 'card-1' });
      expect(JSON.stringify(updates[0].activity)).toContain('Switched conversation');
      expect(connector['state']!.legionConversations['app:alice']).toBe('legion-2');
      expect(connector['state']!.legionConversationTargets['legion-2']).toBe('app:alice');
      expect(connector['state']!.legionConversationTargets['legion-1']).toBeUndefined();
    } finally {
      await connector.stop();
    }
  });

  it('pages the picker without changing state', async () => {
    const { app, deps, updates } = setup();
    const factory = vi.fn<TeamsAppFactory>(() => app);
    const dir = await mkdtemp(join(tmpdir(), 'teams-switch-page-'));
    const connector = new TeamsConnector({ stateDir: dir, transport: 'socket' }, deps, factory);
    const conversations = Array.from({ length: 7 }, (_, index) => ({
      id: `legion-${index}`,
      resolvedTitle: `Conversation ${index}`,
      messageCount: 1,
      updatedAt: `2026-11-${String(20 + index).padStart(2, '0')}T00:00:00.000Z`,
      status: 'active',
    }));
    const callTool = vi.fn(async () => ({
      result: { status: 'success', data: { conversations } },
      conversationId: 'c',
    }));
    await connector.start({ callTool } as never);
    try {
      connector['state']!.legionConversations['app:alice'] = 'legion-0';
      const [entry, handlers] = factory.mock.calls[0];
      await handlers.onConversationAction(entry, {
        send: vi.fn(),
        activity: {
          from: { aadObjectId: 'aad-1' },
          channelData: { tenant: { id: 't' } },
          conversation: { conversationType: 'personal', id: 'conv-1' },
          replyToId: 'card-1',
          value: { action: { data: { action: 'switch', offset: 5, key: 'app:alice' } } },
        },
      } as never);
      await vi.waitFor(() => expect(updates).toHaveLength(1));
      expect(JSON.stringify(updates[0].activity)).toContain('Showing 6–7 of 7');
      expect(connector['state']!.legionConversations['app:alice']).toBe('legion-0');
    } finally {
      await connector.stop();
    }
  });

  it('shows an error card for a conversation outside the active list', async () => {
    const { app, deps, updates } = setup();
    const factory = vi.fn<TeamsAppFactory>(() => app);
    const dir = await mkdtemp(join(tmpdir(), 'teams-switch-missing-'));
    const connector = new TeamsConnector({ stateDir: dir, transport: 'socket' }, deps, factory);
    const callTool = vi.fn(async () => ({
      result: { status: 'success', data: { conversations: [] } },
      conversationId: 'c',
    }));
    await connector.start({ callTool } as never);
    try {
      const [entry, handlers] = factory.mock.calls[0];
      await handlers.onConversationAction(entry, {
        send: vi.fn(),
        activity: {
          from: { aadObjectId: 'aad-1' },
          channelData: { tenant: { id: 't' } },
          conversation: { conversationType: 'personal', id: 'conv-1' },
          replyToId: 'card-1',
          value: { action: { data: {
            action: 'switch', conversationId: 'legion-x', key: 'app:alice',
          } } },
        },
      } as never);
      await vi.waitFor(() => expect(updates).toHaveLength(1));
      expect(JSON.stringify(updates[0].activity)).toContain('Conversation not found.');
    } finally {
      await connector.stop();
    }
  });

  it('rejects an invalid switch payload without listing conversations', async () => {
    const { app, deps, updates } = setup();
    const factory = vi.fn<TeamsAppFactory>(() => app);
    const dir = await mkdtemp(join(tmpdir(), 'teams-switch-invalid-'));
    const connector = new TeamsConnector({ stateDir: dir, transport: 'socket' }, deps, factory);
    const callTool = vi.fn(async () => ({ result: { status: 'success' }, conversationId: 'c' }));
    await connector.start({ callTool } as never);
    try {
      const [entry, handlers] = factory.mock.calls[0];
      const invoke = (data: Record<string, unknown>) => handlers.onConversationAction(entry, {
        send: vi.fn(),
        activity: {
          from: { aadObjectId: 'aad-1' },
          channelData: { tenant: { id: 't' } },
          conversation: { conversationType: 'personal', id: 'conv-1' },
          replyToId: 'card-1',
          value: { action: { data } },
        },
      } as never);
      expect(await invoke({ action: 'switch' })).toHaveProperty('body.value', 'Invalid switch payload.');
      expect(await invoke({ action: 'switch', key: 'app:bob', conversationId: 'c1' }))
        .toHaveProperty('body.value', 'Invalid switch payload.');
      expect(await invoke({ action: 'switch', conversationId: 'x'.repeat(257) }))
        .toHaveProperty('body.value', 'Invalid switch payload.');
      expect(await invoke({ action: 'switch', offset: '5' }))
        .toHaveProperty('body.value', 'Invalid switch payload.');
      expect(callTool).not.toHaveBeenCalled();
      expect(updates).toHaveLength(0);
    } finally {
      await connector.stop();
    }
  });

  it('rejects tenant and conversation-type mismatches for switch', async () => {
    const { app, deps, updates } = setup();
    const factory = vi.fn<TeamsAppFactory>(() => app);
    const dir = await mkdtemp(join(tmpdir(), 'teams-switch-scope-'));
    const connector = new TeamsConnector({ stateDir: dir, transport: 'socket' }, deps, factory);
    const callTool = vi.fn(async () => ({ result: { status: 'success' }, conversationId: 'c' }));
    await connector.start({ callTool } as never);
    try {
      const [entry, handlers] = factory.mock.calls[0];
      const invoke = (tenant: string, conversationType: string) => handlers.onConversationAction(entry, {
        send: vi.fn(),
        activity: {
          from: { aadObjectId: 'aad-1' },
          channelData: { tenant: { id: tenant } },
          conversation: { conversationType, id: 'conv-1' },
          value: { action: { data: { action: 'switch', offset: 0 } } },
        },
      } as never);
      expect(await invoke('other', 'personal'))
        .toHaveProperty('body.value', 'This bot is not available in your tenant.');
      expect(await invoke('t', 'channel'))
        .toHaveProperty('body.value', 'Only personal chats and group chats are supported.');
      expect(callTool).not.toHaveBeenCalled();
      expect(updates).toHaveLength(0);
    } finally {
      await connector.stop();
    }
  });

  it('rejects unregistered personal and group clickers', async () => {
    const { app, deps } = setup();
    const factory = vi.fn<TeamsAppFactory>(() => app);
    const dir = await mkdtemp(join(tmpdir(), 'teams-switch-unknown-'));
    const connector = new TeamsConnector({ stateDir: dir, transport: 'socket' }, deps, factory);
    const callTool = vi.fn(async () => ({ result: { status: 'success' }, conversationId: 'c' }));
    await connector.start({ callTool } as never);
    try {
      const [entry, handlers] = factory.mock.calls[0];
      const invoke = (conversationType: string, id: string) => handlers.onConversationAction(entry, {
        send: vi.fn(),
        activity: {
          from: { aadObjectId: 'aad-9' },
          channelData: { tenant: { id: 't' } },
          conversation: { conversationType, id },
          value: { action: { data: { action: 'switch', offset: 0 } } },
        },
      } as never);
      expect(await invoke('personal', 'conv-1')).toHaveProperty('body.value', 'Not registered.');
      expect(await invoke('groupChat', '19:unknown@thread.v2'))
        .toHaveProperty('body.value', 'This group is not registered.');
      expect(callTool).not.toHaveBeenCalled();
    } finally {
      await connector.stop();
    }
  });

  it('switches a group chat under the group key', async () => {
    const { app, deps, updates } = setup();
    const factory = vi.fn<TeamsAppFactory>(() => app);
    const dir = await mkdtemp(join(tmpdir(), 'teams-switch-group-'));
    const connector = new TeamsConnector({ stateDir: dir, transport: 'socket' }, deps, factory);
    const callTool = vi.fn(async () => ({
      result: { status: 'success', data: { conversations: [{
        id: 'legion-group-2', resolvedTitle: 'Group thread', messageCount: 2,
        updatedAt: '2026-11-25T00:00:00.000Z', status: 'active',
      }] } },
      conversationId: 'c',
    }));
    await connector.start({ callTool } as never);
    try {
      connector['state']!.legionConversations['app:team-eng:19:g@thread.v2'] = 'legion-group-1';
      const [entry, handlers] = factory.mock.calls[0];
      await handlers.onConversationAction(entry, {
        send: vi.fn(),
        activity: {
          from: { aadObjectId: 'aad-1' },
          channelData: { tenant: { id: 't' } },
          conversation: { conversationType: 'groupChat', id: '19:g@thread.v2' },
          replyToId: 'card-1',
          value: { action: { data: {
            action: 'switch', conversationId: 'legion-group-2',
            key: 'app:team-eng:19:g@thread.v2',
          } } },
        },
      } as never);
      await vi.waitFor(() => expect(updates).toHaveLength(1));
      expect(callTool).toHaveBeenCalledWith('team-eng', 'list_conversations', expect.anything());
      expect(connector['state']!.legionConversations['app:team-eng:19:g@thread.v2'])
        .toBe('legion-group-2');
    } finally {
      await connector.stop();
    }
  });

  it('uses the fallback participant for an unknown personal clicker', async () => {
    const { app, deps, participants, updates } = setup();
    participants.push({
      id: 'operator', name: 'Operator', type: 'user', tools: {}, status: 'active',
      identities: [{ connector: 'teams', externalId: 'aad-op' }],
    });
    const factory = vi.fn<TeamsAppFactory>(() => app);
    const dir = await mkdtemp(join(tmpdir(), 'teams-switch-fallback-'));
    const connector = new TeamsConnector({
      stateDir: dir,
      transport: 'socket',
      unknownUserPolicy: 'fallback',
      fallbackUserId: 'operator',
    }, deps, factory);
    const callTool = vi.fn(async () => ({
      result: { status: 'success', data: { conversations: [{
        id: 'legion-op', resolvedTitle: 'Fallback thread', messageCount: 1,
        updatedAt: '2026-11-25T00:00:00.000Z', status: 'active',
      }] } },
      conversationId: 'c',
    }));
    await connector.start({ callTool } as never);
    try {
      const [entry, handlers] = factory.mock.calls[0];
      await handlers.onConversationAction(entry, {
        send: vi.fn(),
        activity: {
          from: { aadObjectId: 'aad-9' },
          channelData: { tenant: { id: 't' } },
          conversation: { conversationType: 'personal', id: 'conv-1' },
          replyToId: 'card-1',
          value: { action: { data: {
            action: 'switch', conversationId: 'legion-op', key: 'app:operator:aad-9',
          } } },
        },
      } as never);
      await vi.waitFor(() => expect(updates).toHaveLength(1));
      expect(callTool).toHaveBeenCalledWith('operator', 'list_conversations', expect.anything());
      expect(connector['state']!.legionConversations['app:operator:aad-9']).toBe('legion-op');
    } finally {
      await connector.stop();
    }
  });

  it('sends a follow-up card when replyToId is absent', async () => {
    const { app, deps, sent } = setup();
    const factory = vi.fn<TeamsAppFactory>(() => app);
    const dir = await mkdtemp(join(tmpdir(), 'teams-switch-followup-'));
    const connector = new TeamsConnector({ stateDir: dir, transport: 'socket' }, deps, factory);
    const callTool = vi.fn(async () => ({
      result: { status: 'success', data: { conversations: [{
        id: 'legion-2', resolvedTitle: 'Second', messageCount: 1,
        updatedAt: '2026-11-25T00:00:00.000Z', status: 'active',
      }] } },
      conversationId: 'c',
    }));
    await connector.start({ callTool } as never);
    try {
      const [entry, handlers] = factory.mock.calls[0];
      await handlers.onConversationAction(entry, {
        send: vi.fn(),
        activity: {
          from: { aadObjectId: 'aad-1' },
          channelData: { tenant: { id: 't' } },
          conversation: { conversationType: 'personal', id: 'conv-1' },
          value: { action: { data: {
            action: 'switch', conversationId: 'legion-2', key: 'app:alice',
          } } },
        },
      } as never);
      await vi.waitFor(() => expect(sent).toHaveLength(1));
      expect(sent[0].conversationId).toBe('conv-1');
      expect(JSON.stringify(sent[0].activity)).toContain('Switched conversation');
    } finally {
      await connector.stop();
    }
  });

  it('serializes a switch behind an in-flight turn for the same chat', async () => {
    const { app, deps, updates } = setup();
    const factory = vi.fn<TeamsAppFactory>(() => app);
    const dir = await mkdtemp(join(tmpdir(), 'teams-switch-queue-'));
    const connector = new TeamsConnector({ stateDir: dir, transport: 'socket' }, deps, factory);
    const callTool = vi.fn(async () => ({
      result: { status: 'success', data: { conversations: [{
        id: 'legion-2', resolvedTitle: 'Second', messageCount: 1,
        updatedAt: '2026-11-25T00:00:00.000Z', status: 'active',
      }] } },
      conversationId: 'c',
    }));
    await connector.start({ callTool } as never);
    try {
      let release!: () => void;
      const gate = new Promise<void>((resolve) => { release = resolve; });
      const turn = connector['queue'].run('app:aad-1', async () => { await gate; });
      const [entry, handlers] = factory.mock.calls[0];
      await handlers.onConversationAction(entry, {
        send: vi.fn(),
        activity: {
          from: { aadObjectId: 'aad-1' },
          channelData: { tenant: { id: 't' } },
          conversation: { conversationType: 'personal', id: 'conv-1' },
          replyToId: 'card-1',
          value: { action: { data: {
            action: 'switch', conversationId: 'legion-2', key: 'app:alice',
          } } },
        },
      } as never);
      expect(updates).toHaveLength(0);
      release();
      await turn;
      await vi.waitFor(() => expect(updates).toHaveLength(1));
    } finally {
      await connector.stop();
    }
  });

});

it('exposes only the configured Teams tools after start', async () => {
  const { app, deps } = setup();
  const dir = await mkdtemp(join(tmpdir(), 'teams-tools-'));
  const connector = new TeamsConnector(
    { stateDir: dir, port: 0, tools: ['cards'] },
    deps,
    () => app,
  );
  await connector.start({} as never);
  try {
    expect(connector.tools().map((tool) => tool.name)).toEqual(['teams__send_card']);
  } finally {
    await connector.stop();
  }
});

it('exposes all Teams tools by default after start', async () => {
  const { app, deps } = setup();
  const dir = await mkdtemp(join(tmpdir(), 'teams-tools-default-'));
  const connector = new TeamsConnector({ stateDir: dir, port: 0 }, deps, () => app);
  await connector.start({} as never);
  try {
    expect(connector.tools().map((tool) => tool.name)).toEqual([
      'teams__send_card',
      'teams__post_reaction',
      'teams__list_conversations',
    ]);
  } finally {
    await connector.stop();
  }
});

it('exposes no tools before start', () => {
  const { app, deps } = setup();
  const connector = new TeamsConnector({ stateDir: '/tmp/x' }, deps, () => app);
  expect(connector.tools()).toEqual([]);
});
