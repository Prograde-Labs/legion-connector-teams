// src/outbound.test.ts
import { describe, expect, it, vi } from 'vitest';
import { deliverMessage } from './outbound.js';
import { emptyState } from './state.js';
import { resolveOptions } from './options.js';
import type { CollectiveLike } from './directory.js';

function setup(participants: Array<Record<string, unknown>>, conversations: Record<string, unknown> = {}) {
  const state = { ...emptyState(), conversations };
  const send = vi.fn(async (_conversationId: string, _activity: unknown) => ({ id: 'act-1' }));
  const create = vi.fn(async () => ({ id: 'conv-new' }));
  const save = vi.fn(async () => undefined);
  const collective = {
    list: () => participants,
    get: (id: string) => participants.find((p) => p.id === id),
  } as unknown as CollectiveLike;
  const apps = {
    ownerOf: (appId: string) => appId === 'app' ? 'assistant' : undefined,
    getByAppId: (appId: string) => (appId === 'app' ? { send, api: { conversations: { create } } } : undefined),
  };
  return { state, send, create, save, collective, apps };
}

describe('deliverMessage', () => {
  it('prefers the exact route and otherwise picks the most recently updated fallback', async () => {
    const s = setup([{ id: 'assistant', identities: [{ connector: 'teams-bot', externalId: 'app' }] }], {
      'app:alice:aad-8': { conversationId: 'older', updatedAt: '2026-11-20T00:00:00Z' },
      'app:alice:aad-9': { conversationId: 'newer', updatedAt: '2026-11-21T00:00:00Z' },
      'app:alice:aad-7': { conversationId: 'oldest', updatedAt: '2026-11-19T00:00:00Z' },
      'app:alice-two:aad-1': { conversationId: 'unrelated', updatedAt: '2026-11-22T00:00:00Z' },
    });
    const deps = { ...s, state: s.state as never, stateStore: { save: s.save } as never, options: resolveOptions({}) };
    const message = { senderId: 'assistant', recipientId: 'alice', conversationId: 'legion-1', content: 'hello' };
    await deliverMessage(deps, message);
    expect(s.send).toHaveBeenLastCalledWith('newer', expect.objectContaining({ text: 'hello' }));
    s.state.conversations['app:alice'] = { conversationId: 'exact', updatedAt: '2026-11-19T00:00:00Z' };
    // A second conversation has no learned target, so it resolves the exact route afresh.
    await deliverMessage(deps, { ...message, conversationId: 'legion-2' });
    expect(s.send).toHaveBeenLastCalledWith('exact', expect.objectContaining({ text: 'hello' }));
    expect(s.create).not.toHaveBeenCalled();
  });
  it('sends approval cards after the message', async () => {
    const s = setup(
      [
        { id: 'assistant', name: 'Assistant', type: 'agent', tools: {},
          identities: [{ connector: 'teams-bot', externalId: 'app' }] },
        { id: 'alice', name: 'Alice', type: 'user', tools: {},
          identities: [{ connector: 'teams', externalId: 'aad-1' }] },
      ],
      { 'app:alice': { conversationId: 'conv-1', updatedAt: 'now' } },
    );
    await deliverMessage(
      { collective: s.collective, state: s.state as never, stateStore: { save: s.save } as never,
        apps: s.apps as never, options: resolveOptions({ stateDir: '/tmp/x', tenantId: 'tenant-1' }) },
      {
        senderId: 'assistant',
        recipientId: 'alice',
        conversationId: 'legion-1',
        content: 'hello',
        approvalRequests: [
          {
            approvalId: 'appr-1',
            conversationId: 'legion-1',
            requesterId: 'assistant',
            tool: 'write_file',
            args: { path: 'notes.txt' },
            createdAt: '2026-01-01T00:00:00.000Z',
          },
        ],
      },
    );
    expect(s.send).toHaveBeenCalledTimes(2);
    expect(JSON.stringify(s.send.mock.calls[1][1])).toContain('Approval requested');
  });

  it('still delivers and saves state when approval card rendering fails', async () => {
    const s = setup(
      [
        { id: 'assistant', name: 'Assistant', type: 'agent', tools: {},
          identities: [{ connector: 'teams-bot', externalId: 'app' }] },
        { id: 'alice', name: 'Alice', type: 'user', tools: {},
          identities: [{ connector: 'teams', externalId: 'aad-1' }] },
      ],
      { 'app:alice': { conversationId: 'conv-1', updatedAt: 'now' } },
    );
    s.send.mockResolvedValueOnce({ id: 'act-1' }).mockRejectedValueOnce(new Error('card down'));
    const log = vi.spyOn(console, 'error').mockImplementation(() => undefined);
    try {
      await expect(
        deliverMessage(
          { collective: s.collective, state: s.state as never, stateStore: { save: s.save } as never,
            apps: s.apps as never, options: resolveOptions({ stateDir: '/tmp/x', tenantId: 'tenant-1' }) },
          {
            senderId: 'assistant',
            recipientId: 'alice',
            conversationId: 'legion-1',
            content: 'hello',
            approvalRequests: [
              {
                approvalId: 'appr-1',
                conversationId: 'legion-1',
                requesterId: 'assistant',
                tool: 'write_file',
                args: {},
                createdAt: '2026-01-01T00:00:00.000Z',
              },
            ],
          },
        ),
      ).resolves.toEqual({ status: 'delivered' });
      expect(s.save).toHaveBeenCalledWith(s.state);
    } finally {
      log.mockRestore();
    }
  });

  it('prefers the latest personal fallback over a newer group record', async () => {
    const s = setup(
      [
        { id: 'assistant', identities: [{ connector: 'teams-bot', externalId: 'app' }] },
        { id: 'operator' },
      ],
      {
        'app:operator:aad-1': {
          conversationId: 'personal-conv', conversationType: 'personal', updatedAt: '2026-11-20T00:00:00Z',
        },
        'app:operator:19:g@thread.v2': {
          conversationId: '19:g@thread.v2', conversationType: 'groupChat', updatedAt: '2026-11-21T00:00:00Z',
        },
      },
    );
    await deliverMessage(
      { collective: s.collective, state: s.state as never, stateStore: { save: s.save } as never,
        apps: s.apps as never, options: resolveOptions({ stateDir: '/tmp/x' }) },
      { senderId: 'assistant', recipientId: 'operator', conversationId: 'legion-1', content: 'hello' },
    );
    expect(s.send).toHaveBeenCalledWith('personal-conv', expect.objectContaining({ text: 'hello' }));
    expect(s.send).not.toHaveBeenCalledWith('19:g@thread.v2', expect.anything());
  });

  it('sends through the sender bot to a stored conversation', async () => {
    const s = setup(
      [
        { id: 'assistant', name: 'Assistant', type: 'agent', tools: {},
          identities: [{ connector: 'teams-bot', externalId: 'app' }] },
        { id: 'alice', name: 'Alice', type: 'user', tools: {},
          identities: [{ connector: 'teams', externalId: 'aad-1' }] },
      ],
      { 'app:alice': { conversationId: 'conv-1', updatedAt: 'now' } },
    );
    await deliverMessage(
      { collective: s.collective, state: s.state as never, stateStore: { save: s.save } as never,
        apps: s.apps as never, options: resolveOptions({ stateDir: '/tmp/x', tenantId: 'tenant-1' }) },
      { senderId: 'assistant', recipientId: 'alice', conversationId: 'legion-1', content: 'hello' },
    );
    expect(s.send).toHaveBeenCalledWith('conv-1', expect.anything());
    expect(s.create).not.toHaveBeenCalled();
  });

  it('creates a personal conversation for a resolved recipient instead of using a stale fallback', async () => {
    const s = setup([
      { id: 'assistant', name: 'Assistant', type: 'agent', tools: {},
        identities: [{ connector: 'teams-bot', externalId: 'app' }] },
      { id: 'alice', name: 'Alice', type: 'user', tools: {},
        identities: [{ connector: 'teams', externalId: 'aad-1' }] },
    ], { 'app:alice:aad-9': { conversationId: 'stale', updatedAt: '2026-11-21T00:00:00Z' } });
    await deliverMessage(
      { collective: s.collective, state: s.state as never, stateStore: { save: s.save } as never,
        apps: s.apps as never, options: resolveOptions({ stateDir: '/tmp/x', tenantId: 'tenant-1' }) },
      { senderId: 'assistant', recipientId: 'alice', conversationId: 'legion-1', content: 'hello' },
    );
    expect(s.create).toHaveBeenCalledWith(expect.objectContaining({
      members: [{ id: 'aad-1' }],
      tenantId: 'tenant-1',
      channelData: { tenant: { id: 'tenant-1' } },
    }));
    expect(s.state.conversations['app:alice']).toEqual(expect.objectContaining({
      conversationId: 'conv-new',
      conversationType: 'personal',
      externalId: 'aad-1',
      tenantId: 'tenant-1',
    }));
    expect(s.state.legionConversationTargets['legion-1']).toBe('app:alice');
    expect(s.save).toHaveBeenCalledWith(s.state);
    expect(s.send.mock.calls).toEqual([['conv-new', expect.objectContaining({ text: 'hello' })]]);
  });

  it('reports not-reachable when the sender has no Teams bot', async () => {
    const s = setup([
      { id: 'assistant', name: 'Assistant', type: 'agent', tools: {} },
      { id: 'alice', name: 'Alice', type: 'user', tools: {},
        identities: [{ connector: 'teams', externalId: 'aad-1' }] },
    ]);
    await expect(deliverMessage(
      { collective: s.collective, state: s.state as never, stateStore: { save: s.save } as never,
        apps: s.apps as never, options: resolveOptions({ stateDir: '/tmp/x' }) },
      { senderId: 'assistant', recipientId: 'alice', conversationId: 'legion-1', content: 'hello' },
    )).resolves.toEqual({
      status: 'not-reachable',
      reason: "sender 'assistant' has no live Teams bot",
    });
    expect(s.send).not.toHaveBeenCalled();
  });

  it('falls back to the replyTo participant bot when the sender has no Teams bot', async () => {
    const s = setup(
      [
        { id: 'assistant', name: 'Assistant', type: 'agent', tools: {},
          identities: [{ connector: 'teams-bot', externalId: 'app' }] },
        { id: 'quiet', name: 'Quiet', type: 'service', tools: {} },
        { id: 'alice', name: 'Alice', type: 'user', tools: {},
          identities: [{ connector: 'teams', externalId: 'aad-1' }] },
      ],
      { 'app:alice': { conversationId: 'conv-1', updatedAt: 'now' } },
    );
    await deliverMessage(
      { collective: s.collective, state: s.state as never, stateStore: { save: s.save } as never,
        apps: s.apps as never, options: resolveOptions({ stateDir: '/tmp/x', tenantId: 'tenant-1' }) },
      {
        senderId: 'quiet',
        recipientId: 'alice',
        conversationId: 'legion-1',
        replyTo: 'assistant',
        content: 'hello',
      },
    );
    expect(s.send).toHaveBeenCalledWith('conv-1', expect.objectContaining({ text: 'hello' }));
  });

  it('reports failed and logs a sanitized Teams send error', async () => {
    const s = setup([
      { id: 'assistant', identities: [{ connector: 'teams-bot', externalId: 'app' }] },
    ], { 'app:alice': { conversationId: 'conv-1', updatedAt: 'now' } });
    const error = Object.assign(new Error('403 Bearer abc123'), { authorization: 'Bearer secret-field' });
    s.send.mockRejectedValueOnce(error);
    const log = vi.spyOn(console, 'error').mockImplementation(() => undefined);
    try {
      await expect(deliverMessage(
        { collective: s.collective, state: s.state as never, stateStore: { save: s.save } as never,
          apps: s.apps as never, options: resolveOptions({ stateDir: '/tmp/x' }) },
        { senderId: 'assistant', recipientId: 'alice', conversationId: 'legion-1', content: 'hello' },
      )).resolves.toEqual({
        status: 'failed',
        reason: 'Teams message delivery failed: Error: 403 Bearer [redacted]',
      });
      expect(log).toHaveBeenCalledWith('Teams message delivery failed', 'Error: 403 Bearer [redacted]');
      expect(JSON.stringify(log.mock.calls)).not.toContain('abc123');
      expect(JSON.stringify(log.mock.calls)).not.toContain('secret-field');
    } finally {
      log.mockRestore();
    }
  });

  it('reports failed and logs a sanitized creation error without saving or sending', async () => {
    const s = setup([
      { id: 'assistant', identities: [{ connector: 'teams-bot', externalId: 'app' }] },
      { id: 'alice', identities: [{ connector: 'teams', externalId: 'aad-1' }] },
    ]);
    s.create.mockRejectedValueOnce(new Error('403 Bearer abc123'));
    const log = vi.spyOn(console, 'error').mockImplementation(() => undefined);
    try {
      await expect(deliverMessage(
        { collective: s.collective, state: s.state as never, stateStore: { save: s.save } as never,
          apps: s.apps as never, options: resolveOptions({ stateDir: '/tmp/x' }) },
        { senderId: 'assistant', recipientId: 'alice', conversationId: 'legion-1', content: 'hello' },
      )).resolves.toEqual({
        status: 'failed',
        reason: 'Teams conversation creation failed: Error: 403 Bearer [redacted]',
      });
      expect(log).toHaveBeenCalledWith('Teams conversation creation failed', 'Error: 403 Bearer [redacted]');
      expect(JSON.stringify(log.mock.calls)).not.toContain('abc123');
      expect(s.state.conversations).toEqual({});
      expect(s.save).not.toHaveBeenCalled();
      expect(s.send).not.toHaveBeenCalled();
    } finally {
      log.mockRestore();
    }
  });

  it('reports not-reachable when the bot app or recipient identity is missing', async () => {
    const s = setup([
      { id: 'assistant', identities: [{ connector: 'teams-bot', externalId: 'other' }] },
      { id: 'alice' },
    ]);
    const deps = { collective: s.collective, state: s.state as never, stateStore: { save: s.save } as never,
      apps: s.apps as never, options: resolveOptions({ stateDir: '/tmp/x' }) };
    const message = { senderId: 'assistant', recipientId: 'alice', conversationId: 'legion-1', content: 'hello' };
    await expect(deliverMessage(deps, message)).resolves.toEqual({
      status: 'not-reachable',
      reason: "sender 'assistant' has no live Teams bot",
    });
    expect(s.create).not.toHaveBeenCalled();
    (s.collective.get('assistant')!.identities![0]).externalId = 'app';
    await expect(deliverMessage(deps, message)).resolves.toEqual({
      status: 'not-reachable',
      reason: "no Teams conversation for 'alice'",
    });
    expect(s.create).not.toHaveBeenCalled();
    expect(s.send).not.toHaveBeenCalled();
  });

  it('sets the target map and learns the thread by default', async () => {
    const s = setup([
      { id: 'assistant', identities: [{ connector: 'teams-bot', externalId: 'app' }] },
      { id: 'alice', identities: [{ connector: 'teams', externalId: 'aad-1' }] },
    ], { 'app:alice': { conversationId: 'conv-1', updatedAt: 'now' } });
    await expect(deliverMessage(
      { collective: s.collective, state: s.state as never, stateStore: { save: s.save } as never,
        apps: s.apps as never, options: resolveOptions({ stateDir: '/tmp/x' }) },
      { senderId: 'assistant', recipientId: 'alice', conversationId: 'legion-1', content: 'hello' },
    )).resolves.toEqual({ status: 'delivered' });
    expect(s.state.legionConversationTargets['legion-1']).toBe('app:alice');
    expect(s.state.legionConversations['app:alice']).toBe('legion-1');
    expect(s.save).toHaveBeenCalledWith(s.state);
  });

  it('sets the target map without changing the thread when learning is off', async () => {
    const s = setup([
      { id: 'assistant', identities: [{ connector: 'teams-bot', externalId: 'app' }] },
      { id: 'alice', identities: [{ connector: 'teams', externalId: 'aad-1' }] },
    ], { 'app:alice': { conversationId: 'conv-1', updatedAt: 'now' } });
    s.state.legionConversations['app:alice'] = 'legion-previous';
    await expect(deliverMessage(
      { collective: s.collective, state: s.state as never, stateStore: { save: s.save } as never,
        apps: s.apps as never,
        options: resolveOptions({ stateDir: '/tmp/x', learnOutboundThread: false }) },
      { senderId: 'assistant', recipientId: 'alice', conversationId: 'legion-1', content: 'hello' },
    )).resolves.toEqual({ status: 'delivered' });
    expect(s.state.legionConversationTargets['legion-1']).toBe('app:alice');
    expect(s.state.legionConversations['app:alice']).toBe('legion-previous');
    expect(s.save).toHaveBeenCalledWith(s.state);
  });

  it('learns the group thread by default', async () => {
    const s = setup([
      { id: 'assistant', identities: [{ connector: 'teams-bot', externalId: 'app' }] },
      { id: 'team-eng', identities: [{ connector: 'teams-group', externalId: '19:g@thread.v2' }] },
    ]);
    await expect(deliverMessage(
      { collective: s.collective, state: s.state as never, stateStore: { save: s.save } as never,
        apps: s.apps as never, options: resolveOptions({ stateDir: '/tmp/x' }) },
      { senderId: 'assistant', recipientId: 'team-eng', conversationId: 'legion-1', content: 'hello' },
    )).resolves.toEqual({ status: 'delivered' });
    expect(s.state.legionConversationTargets['legion-1']).toBe('app:team-eng:19:g@thread.v2');
    expect(s.state.legionConversations['app:team-eng:19:g@thread.v2']).toBe('legion-1');
  });

  it('sends chunks with the configured AI label', async () => {
    const s = setup([
      { id: 'assistant', identities: [{ connector: 'teams-bot', externalId: 'app' }] },
    ], { 'app:alice': { conversationId: 'conv-1', updatedAt: 'now' } });
    await deliverMessage(
      { collective: s.collective, state: s.state as never, stateStore: { save: s.save } as never,
        apps: s.apps as never, options: resolveOptions({ stateDir: '/tmp/x', maxMessageChars: 3, aiLabel: false }) },
      { senderId: 'assistant', recipientId: 'alice', conversationId: 'legion-1', content: 'abcdef' },
    );
    expect(s.send).toHaveBeenCalledTimes(2);
    expect(s.send.mock.calls.map((call: unknown[]) => call[1])).toEqual([
      expect.objectContaining({ text: 'abc' }), expect.objectContaining({ text: 'def' }),
    ]);
    expect(s.send.mock.calls[0][1]).toHaveProperty('entities', undefined);
    s.send.mockClear();
    await deliverMessage(
      { collective: s.collective, state: s.state as never, stateStore: { save: s.save } as never,
        apps: s.apps as never, options: resolveOptions({ stateDir: '/tmp/x', aiLabel: true }) },
      { senderId: 'assistant', recipientId: 'alice', conversationId: 'legion-1', content: 'hello' },
    );
    expect(s.send.mock.calls[0][1]).toHaveProperty('entities', [
      expect.objectContaining({ additionalType: ['AIGeneratedContent'] }),
    ]);
  });

  it('routes a Legion conversation to its target group', async () => {
    const s = setup(
      [
        { id: 'assistant', identities: [{ connector: 'teams-bot', externalId: 'app' }] },
        { id: 'team-eng', identities: [{ connector: 'teams-group', externalId: '19:other@thread.v2' }] },
      ],
      {
        'app:team-eng:19:g@thread.v2': { conversationId: '19:g@thread.v2', updatedAt: '2026-11-21T00:00:00Z' },
        'app:team-eng:19:other@thread.v2': { conversationId: '19:other@thread.v2', updatedAt: '2026-11-22T00:00:00Z' },
      },
    );
    s.state.legionConversationTargets['legion-1'] = 'app:team-eng:19:g@thread.v2';
    await deliverMessage(
      { collective: s.collective, state: s.state as never, stateStore: { save: s.save } as never,
        apps: s.apps as never, options: resolveOptions({ stateDir: '/tmp/x' }) },
      { senderId: 'assistant', recipientId: 'team-eng', conversationId: 'legion-1', content: 'hello' },
    );
    expect(s.send).toHaveBeenCalledWith('19:g@thread.v2', expect.objectContaining({ text: 'hello' }));
    expect(s.send).not.toHaveBeenCalledWith('19:other@thread.v2', expect.anything());
  });

  it('sends directly to a teams-group identity and records the target', async () => {
    const s = setup([
      { id: 'assistant', identities: [{ connector: 'teams-bot', externalId: 'app' }] },
      { id: 'team-eng', identities: [{ connector: 'teams-group', externalId: '19:g@thread.v2' }] },
    ]);
    await deliverMessage(
      { collective: s.collective, state: s.state as never, stateStore: { save: s.save } as never,
        apps: s.apps as never, options: resolveOptions({ stateDir: '/tmp/x' }) },
      { senderId: 'assistant', recipientId: 'team-eng', conversationId: 'legion-1', content: 'hello' },
    );
    expect(s.send).toHaveBeenCalledWith('19:g@thread.v2', expect.objectContaining({ text: 'hello' }));
    expect(s.state.conversations['app:team-eng:19:g@thread.v2']).toMatchObject({
      conversationId: '19:g@thread.v2',
      conversationType: 'groupChat',
      externalId: '19:g@thread.v2',
    });
    expect(s.state.legionConversationTargets['legion-1']).toBe('app:team-eng:19:g@thread.v2');
    expect(s.save).toHaveBeenCalledWith(s.state);
    expect(s.create).not.toHaveBeenCalled();
  });

  it('keeps fallback group replies in the originating group', async () => {
    const s = setup(
      [
        { id: 'assistant', identities: [{ connector: 'teams-bot', externalId: 'app' }] },
        { id: 'operator' },
      ],
      {
        'app:operator:19:g1@thread.v2': { conversationId: '19:g1@thread.v2', updatedAt: '2026-11-21T00:00:00Z' },
        'app:operator:19:g2@thread.v2': { conversationId: '19:g2@thread.v2', updatedAt: '2026-11-22T00:00:00Z' },
      },
    );
    s.state.legionConversationTargets['legion-1'] = 'app:operator:19:g1@thread.v2';
    await deliverMessage(
      { collective: s.collective, state: s.state as never, stateStore: { save: s.save } as never,
        apps: s.apps as never, options: resolveOptions({ stateDir: '/tmp/x' }) },
      { senderId: 'assistant', recipientId: 'operator', conversationId: 'legion-1', content: 'hello' },
    );
    expect(s.send).toHaveBeenCalledWith('19:g1@thread.v2', expect.objectContaining({ text: 'hello' }));
    expect(s.send).not.toHaveBeenCalledWith('19:g2@thread.v2', expect.anything());
  });
  it('does not reuse a learned target for a different recipient', async () => {
    const s = setup([
      { id: 'assistant', identities: [{ connector: 'teams-bot', externalId: 'app' }] },
      { id: 'alice', identities: [{ connector: 'teams', externalId: 'aad-a' }] },
      { id: 'bob', identities: [{ connector: 'teams', externalId: 'aad-b' }] },
    ], { 'app:alice': { conversationId: 'alice-chat', updatedAt: 'now' } });
    const deps = { ...s, state: s.state as never, stateStore: { save: s.save } as never, options: resolveOptions({}) };
    const message = { senderId: 'assistant', recipientId: 'alice', conversationId: 'legion-1', content: 'hello' };
    await deliverMessage(deps, message);
    s.send.mockClear();
    await expect(deliverMessage(deps, { ...message, recipientId: 'bob' })).resolves.toEqual({ status: 'delivered' });
    expect(s.create).toHaveBeenCalledWith(expect.objectContaining({ members: [{ id: 'aad-b' }] }));
    expect(s.send).toHaveBeenCalledTimes(1);
    expect(s.send).toHaveBeenCalledWith('conv-new', expect.anything());
  });

  it('reports delivered if saving state fails after a successful send', async () => {
    const s = setup([{ id: 'assistant', identities: [{ connector: 'teams-bot', externalId: 'app' }] }],
      { 'app:alice': { conversationId: 'alice-chat', updatedAt: 'now' } });
    s.save.mockRejectedValueOnce(new Error('disk error'));
    const log = vi.spyOn(console, 'error').mockImplementation(() => undefined);
    try {
      await expect(deliverMessage(
        { ...s, state: s.state as never, stateStore: { save: s.save } as never, options: resolveOptions({}) },
        { senderId: 'assistant', recipientId: 'alice', conversationId: 'legion-1', content: 'hello' },
      )).resolves.toEqual({ status: 'delivered' });
      expect(s.send).toHaveBeenCalledTimes(1);
      expect(log).toHaveBeenCalledWith('Teams state save failed after delivery', expect.stringContaining('disk error'));
    } finally { log.mockRestore(); }
  });

});

describe('deliverMessage message records', () => {
  it('records outbound activity ids and maps the first chunk to the Legion message', async () => {
    const s = setup(
      [
        { id: 'assistant', identities: [{ connector: 'teams-bot', externalId: 'app' }] },
        { id: 'alice', identities: [{ connector: 'teams', externalId: 'aad-1' }] },
      ],
      { 'app:alice': { conversationId: 'conv-1', updatedAt: 'now' } },
    );
    const send = vi.fn(async () => ({ id: 'teams-act-1' }));
    s.apps.getByAppId = () => ({ send, api: { conversations: { create: s.create } } }) as never;

    await deliverMessage(
      {
        collective: s.collective,
        state: s.state as never,
        stateStore: { save: s.save } as never,
        apps: s.apps as never,
        options: resolveOptions({ stateDir: '/tmp/x', maxMessageRecords: 10 }),
      },
      {
        id: 'legion-msg-1',
        senderId: 'assistant',
        recipientId: 'alice',
        conversationId: 'legion-1',
        content: 'hello',
      },
    );

    expect(s.state.messages['app:teams-act-1']).toMatchObject({
      direction: 'outbound',
      conversationId: 'conv-1',
      legionMessageId: 'legion-msg-1',
      legionConversationId: 'legion-1',
    });
    expect(s.state.legionMessages['legion-msg-1']).toBe('app:teams-act-1');
  });
});
