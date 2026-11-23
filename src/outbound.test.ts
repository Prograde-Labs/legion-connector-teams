// src/outbound.test.ts
import { describe, expect, it, vi } from 'vitest';
import { deliverMessage } from './outbound.js';
import { emptyState } from './state.js';
import { resolveOptions } from './options.js';
import type { CollectiveLike } from './directory.js';

function setup(participants: Array<Record<string, unknown>>, conversations: Record<string, unknown> = {}) {
  const state = { ...emptyState(), conversations };
  const send = vi.fn(async (_conversationId: string, _activity: unknown) => undefined);
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
    const message = { senderId: 'assistant', recipientId: 'alice', content: 'hello' };
    await deliverMessage(deps, message);
    expect(s.send).toHaveBeenLastCalledWith('newer', expect.objectContaining({ text: 'hello' }));
    s.state.conversations['app:alice'] = { conversationId: 'exact', updatedAt: '2026-11-19T00:00:00Z' };
    await deliverMessage(deps, message);
    expect(s.send).toHaveBeenLastCalledWith('exact', expect.objectContaining({ text: 'hello' }));
    expect(s.create).not.toHaveBeenCalled();
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
      { senderId: 'assistant', recipientId: 'alice', content: 'hello' },
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
      { senderId: 'assistant', recipientId: 'alice', content: 'hello' },
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
    expect(s.save).toHaveBeenCalledWith(s.state);
    expect(s.send.mock.calls).toEqual([['conv-new', expect.objectContaining({ text: 'hello' })]]);
  });

  it('drops when the sender has no Teams bot', async () => {
    const s = setup([
      { id: 'assistant', name: 'Assistant', type: 'agent', tools: {} },
      { id: 'alice', name: 'Alice', type: 'user', tools: {},
        identities: [{ connector: 'teams', externalId: 'aad-1' }] },
    ]);
    await deliverMessage(
      { collective: s.collective, state: s.state as never, stateStore: { save: s.save } as never,
        apps: s.apps as never, options: resolveOptions({ stateDir: '/tmp/x' }) },
      { senderId: 'assistant', recipientId: 'alice', content: 'hello' },
    );
    expect(s.send).not.toHaveBeenCalled();
  });

  it('logs and drops a failed Teams send', async () => {
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
        { senderId: 'assistant', recipientId: 'alice', content: 'hello' },
      )).resolves.toBeUndefined();
      expect(log).toHaveBeenCalledWith('Teams message delivery failed', 'Error: 403 Bearer [redacted]');
      expect(JSON.stringify(log.mock.calls)).not.toContain('abc123');
      expect(JSON.stringify(log.mock.calls)).not.toContain('secret-field');
    } finally {
      log.mockRestore();
    }
  });

  it('logs a sanitized creation failure without saving or sending', async () => {
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
        { senderId: 'assistant', recipientId: 'alice', content: 'hello' },
      )).resolves.toBeUndefined();
      expect(log).toHaveBeenCalledWith('Teams conversation creation failed', 'Error: 403 Bearer [redacted]');
      expect(JSON.stringify(log.mock.calls)).not.toContain('abc123');
      expect(s.state.conversations).toEqual({});
      expect(s.save).not.toHaveBeenCalled();
      expect(s.send).not.toHaveBeenCalled();
    } finally {
      log.mockRestore();
    }
  });

  it('drops when the bot app or recipient identity is missing', async () => {
    const s = setup([
      { id: 'assistant', identities: [{ connector: 'teams-bot', externalId: 'other' }] },
      { id: 'alice' },
    ]);
    const deps = { collective: s.collective, state: s.state as never, stateStore: { save: s.save } as never,
      apps: s.apps as never, options: resolveOptions({ stateDir: '/tmp/x' }) };
    const message = { senderId: 'assistant', recipientId: 'alice', content: 'hello' };
    await deliverMessage(deps, message);
    expect(s.create).not.toHaveBeenCalled();
    (s.collective.get('assistant')!.identities![0]).externalId = 'app';
    await deliverMessage(deps, message);
    expect(s.create).not.toHaveBeenCalled();
    expect(s.send).not.toHaveBeenCalled();
  });

  it('sends chunks with the configured AI label', async () => {
    const s = setup([
      { id: 'assistant', identities: [{ connector: 'teams-bot', externalId: 'app' }] },
    ], { 'app:alice': { conversationId: 'conv-1', updatedAt: 'now' } });
    await deliverMessage(
      { collective: s.collective, state: s.state as never, stateStore: { save: s.save } as never,
        apps: s.apps as never, options: resolveOptions({ stateDir: '/tmp/x', maxMessageChars: 3, aiLabel: false }) },
      { senderId: 'assistant', recipientId: 'alice', content: 'abcdef' },
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
      { senderId: 'assistant', recipientId: 'alice', content: 'hello' },
    );
    expect(s.send.mock.calls[0][1]).toHaveProperty('entities', [
      expect.objectContaining({ additionalType: ['AIGeneratedContent'] }),
    ]);
  });
});
