// src/inbound.test.ts
import { describe, expect, it, vi } from 'vitest';
import { handleInbound } from './inbound.js';
import { emptyState } from './state.js';
import { resolveOptions } from './options.js';
import type { BotEntry, CollectiveLike } from './directory.js';
import type { ConnectorContext, StreamChunk } from '@legion-collective/core';
import { rememberPendingPair } from './pairing.js';

function setup(participants: Array<Record<string, unknown>> = []) {
  const state = emptyState();
  const save = vi.fn(async () => undefined);
  const callTool = vi.fn(async (): Promise<{ result: { status: string; error?: string }; conversationId: string }> =>
    ({ result: { status: 'success' }, conversationId: 'c' }));
  const streamTool = vi.fn(async () => ({
    conversationId: 'conv-1',
    gen: (async function* (): AsyncGenerator<StreamChunk> {
      yield { type: 'text_delta', delta: 'Hi' };
      yield { type: 'stream:done', result: { status: 'success', data: { conversationId: 'thread-1' } } };
    })(),
  }));
  const collective = {
    list: () => participants,
    get: (id: string) => participants.find((p) => p.id === id),
    findByIdentity: vi.fn((connector: string, externalId: string) =>
      participants.find((p) =>
        (p.identities as Array<{ connector: string; externalId: string }> | undefined)?.some(
          (i) => i.connector === connector && i.externalId === externalId,
        ),
      )),
  } as unknown as CollectiveLike;
  const bot: BotEntry = {
    participantId: 'assistant', name: 'Assistant', appId: 'app',
    credentials: { appId: 'app', clientSecret: 's', tenantId: 't' },
  };
  return {
    state, save, callTool, streamTool, collective, bot,
    ctx: { callTool, streamTool, registry: { setActive: vi.fn() } } as unknown as ConnectorContext,
  };
}

const activity = (text: string, aad = 'aad-1') => ({
  from: { id: '29:1', name: 'Alice', aadObjectId: aad },
  conversation: { id: 'conv-1', conversationType: 'personal' },
  channelData: { tenant: { id: 't' } },
  serviceUrl: 'https://smba.trafficmanager.net/teams/',
  text,
});

const captureReply = () => vi.fn(async (_value: unknown) => undefined);
const label = { entities: [expect.objectContaining({ additionalType: ['AIGeneratedContent'] })] };
const alice = { id: 'alice', identities: [{ connector: 'teams', externalId: 'aad-1' }] };

function inboundDeps(s: ReturnType<typeof setup>, options = {}) {
  return { ...s, stateStore: { save: s.save } as never,
    options: resolveOptions({ stateDir: '/tmp/x', typingIntervalMs: 0, ...options }) };
}

describe('handleInbound', () => {
  it.each([
    [{ conversation: { conversationType: 'groupChat' } }, 'Only personal chats are supported.'],
    [{ conversation: { id: 'conv-1' } }, 'Only personal chats are supported.'],
    [{ channelData: { tenant: { id: 'other' } } }, 'This bot is not available in your tenant.'],
    [{ channelData: undefined }, 'This bot is not available in your tenant.'],
  ] as const)('rejects unsupported chat or tenant data before resolving identity', async (override, text) => {
    const s = setup([alice]);
    const send = captureReply();
    await handleInbound(inboundDeps(s), { activity: { ...activity('hello'), ...override }, send });
    expect(send).toHaveBeenCalledOnce();
    expect(send).toHaveBeenCalledWith(expect.objectContaining({ text }));
    expect(s.collective.findByIdentity).not.toHaveBeenCalled();
    expect(s.save).not.toHaveBeenCalled();
  });

  it('lets a registered actor redeem a pending user’s code', async () => {
    const s = setup([alice]);
    const send = captureReply();
    const deps = inboundDeps(s);
    await handleInbound(deps, { activity: activity('hello', 'aad-9'), send });
    const [code] = Object.keys(s.state.pendingPairs);
    await handleInbound(deps, { activity: activity(`/pair ${code} bob`), send });
    expect(s.callTool).toHaveBeenCalledOnce();
    expect(s.callTool).toHaveBeenCalledWith('alice', 'pair', {
      participantId: 'bob', connector: 'teams', externalId: 'aad-9',
    });
    expect(s.state.pendingPairs[code]).toBeUndefined();
    expect(send).toHaveBeenLastCalledWith(expect.objectContaining({ text: expect.stringContaining('Paired bob'), ...label }));
  });

  it('isolates fallback senders and does not let them redeem pairing codes', async () => {
    const s = setup([alice]);
    const send = captureReply();
    const deps = inboundDeps(s, { unknownUserPolicy: 'fallback', fallbackUserId: 'alice' });
    for (const aad of ['aad-8', 'aad-9']) {
      await handleInbound(deps, { activity: { ...activity('hello', aad), conversation: {
        id: `conv-${aad}`, conversationType: 'personal',
      } }, send });
    }
    expect(s.state.conversations).toMatchObject({
      'app:alice:aad-8': { conversationId: 'conv-aad-8' },
      'app:alice:aad-9': { conversationId: 'conv-aad-9' },
    });
    expect(s.state.legionConversations).toEqual({ 'app:alice:aad-8': 'thread-1', 'app:alice:aad-9': 'thread-1' });
    expect(s.streamTool).toHaveBeenNthCalledWith(2, 'alice', 'communicate', expect.objectContaining({ conversationId: undefined }));
    await handleInbound(deps, { activity: activity('/new', 'aad-8'), send });
    expect(s.state.legionConversations).toEqual({ 'app:alice:aad-9': 'thread-1' });
    const code = rememberPendingPair(
      s.state, { appId: 'app', connector: 'teams', externalId: 'aad-7' }, 600,
    );
    await handleInbound(deps, { activity: activity(`/pair ${code} bob`, 'aad-8'), send });
    expect(s.callTool).not.toHaveBeenCalled();
    expect(s.state.pendingPairs[code]).toBeDefined();
  });

  it('streams communicate for a resolved user and stores state', async () => {
    const s = setup([
      { id: 'alice', name: 'Alice', type: 'user', tools: {}, status: 'active',
        identities: [{ connector: 'teams', externalId: 'aad-1' }] },
    ]);
    const send = vi.fn(async (_value: unknown) => undefined);
    await handleInbound(
      { bot: s.bot, ctx: s.ctx, collective: s.collective, state: s.state,
        stateStore: { save: s.save } as never, options: resolveOptions({ stateDir: '/tmp/x' }) },
      { activity: activity('hello'), send },
    );
    expect(s.streamTool).toHaveBeenCalledWith('alice', 'communicate', expect.objectContaining({
      to: 'assistant', message: 'hello',
    }));
    expect(s.state.legionConversations['app:alice']).toBe('thread-1');
    await handleInbound(inboundDeps(s), { activity: activity('second turn'), send });
    expect(s.streamTool).toHaveBeenNthCalledWith(2, 'alice', 'communicate', expect.objectContaining({ conversationId: 'thread-1' }));
    expect(s.state.conversations['app:alice'].conversationId).toBe('conv-1');
    expect(s.save).toHaveBeenCalled();
  });

  it('persists the returned thread even when the tool card edit fails', async () => {
    const s = setup([alice]);
    s.streamTool.mockResolvedValueOnce({
      conversationId: 'outer',
      gen: (async function* (): AsyncGenerator<StreamChunk> {
        yield { type: 'tool_call_start', index: 0, id: 't1', name: 'search' };
        yield { type: 'tool_call_args_delta', index: 0, delta: 'query' };
        yield { type: 'stream:done', result: { status: 'success', data: { conversationId: 'thread-2' } } };
      })(),
    });
    const send = vi.fn(async (value: unknown) => {
      if ((value as { id?: string }).id) throw new Error('card edit failed');
      return { id: 'card-1' };
    });
    const log = vi.spyOn(console, 'error').mockImplementation(() => {});
    try {
      await handleInbound(inboundDeps(s, { toolDisplay: 'cards' }), { activity: activity('hello'), send });
      expect(s.save).toHaveBeenCalledTimes(2);
      expect(s.save).toHaveBeenLastCalledWith(expect.objectContaining({ legionConversations: { 'app:alice': 'thread-2' } }));
    } finally { log.mockRestore(); }
  });

  it('keeps the stored thread when an error result lacks a thread id', async () => {
    const s = setup([alice]);
    s.state.legionConversations['app:alice'] = 'old-thread';
    s.streamTool.mockResolvedValueOnce({
      conversationId: 'outer',
      gen: (async function* (): AsyncGenerator<StreamChunk> {
        yield { type: 'stream:done', result: { status: 'error', error: 'failed' } };
      })(),
    });
    await handleInbound(inboundDeps(s), { activity: activity('hello'), send: captureReply() });
    expect(s.state.legionConversations['app:alice']).toBe('old-thread');
  });

  it('issues a pairing code for an unknown user', async () => {
    const s = setup();
    let release!: () => void;
    s.save.mockImplementationOnce(() => new Promise<undefined>((resolve) => { release = () => resolve(undefined); }));
    const send = vi.fn(async (_value: unknown) => {
      expect(s.save).toHaveBeenCalledOnce();
      expect(Object.keys(s.state.pendingPairs)).toHaveLength(1);
    });
    const handled = handleInbound(
      { bot: s.bot, ctx: s.ctx, collective: s.collective, state: s.state,
        stateStore: { save: s.save } as never, options: resolveOptions({ stateDir: '/tmp/x' }) },
      { activity: activity('hello', 'aad-9'), send },
    );
    expect(send).not.toHaveBeenCalled();
    release();
    await handled;
    expect(send).toHaveBeenCalledWith(expect.objectContaining({ text: expect.stringMatching(/Not registered/), ...label }));
    expect(Object.keys(s.state.pendingPairs)).toHaveLength(1);
  });

  it('handles /new without calling communicate', async () => {
    const s = setup([
      { id: 'alice', name: 'Alice', type: 'user', tools: {}, status: 'active',
        identities: [{ connector: 'teams', externalId: 'aad-1' }] },
    ]);
    s.state.legionConversations['app:alice'] = 'old';
    const send = vi.fn(async (_value: unknown) => undefined);
    await handleInbound(
      { bot: s.bot, ctx: s.ctx, collective: s.collective, state: s.state,
        stateStore: { save: s.save } as never, options: resolveOptions({ stateDir: '/tmp/x' }) },
      { activity: activity('/new'), send },
    );
    expect(s.state.legionConversations['app:alice']).toBeUndefined();
    expect(s.streamTool).not.toHaveBeenCalled();
  });

  it('auto-provisions then pairs an unknown user', async () => {
    const s = setup();
    const send = vi.fn(async (_value: unknown) => undefined);
    await handleInbound(
      { bot: s.bot, ctx: s.ctx, collective: s.collective, state: s.state,
        stateStore: { save: s.save } as never,
        options: resolveOptions({ stateDir: '/tmp/x', unknownUserPolicy: 'auto-provision' }) },
      { activity: activity('hello', 'aad-7'), send },
    );
    expect(s.callTool).toHaveBeenCalledWith('operator', 'create_user', expect.objectContaining({
      name: 'Alice',
      tools: { communicate: 'auto', approval_response: 'auto' },
    }));
    expect(s.callTool).toHaveBeenCalledWith('operator', 'pair', expect.objectContaining({
      connector: 'teams', externalId: 'aad-7',
    }));
    expect(send.mock.calls[0][0]).toHaveProperty('text', expect.stringMatching(/Welcome/));
  });

  it('tells rejected users how to obtain access', async () => {
    const s = setup();
    const send = vi.fn(async (_value: unknown) => undefined);
    await handleInbound(
      { ...s, stateStore: { save: s.save } as never,
        options: resolveOptions({ stateDir: '/tmp/x', unknownUserPolicy: 'reject' }) },
      { activity: activity('hello'), send },
    );
    expect(send.mock.calls[0][0]).toHaveProperty('text', expect.stringMatching(/aad-1.*administrator.*register/i));
  });

  it.each(['', '   '])('falls back to the Teams ID for blank AAD ID %j', async (aad) => {
    const s = setup([{ id: 'alice', name: 'Alice', type: 'user', tools: {}, status: 'active',
      identities: [{ connector: 'teams', externalId: '29:1' }] }]);
    const send = vi.fn(async (_value: unknown) => undefined);
    await handleInbound(
      { ...s, stateStore: { save: s.save } as never, options: resolveOptions({ stateDir: '/tmp/x' }) },
      { activity: activity('hello', aad), send },
    );
    expect(s.streamTool).toHaveBeenCalledWith('alice', 'communicate', expect.anything());
  });

  it.each([
    ['create_user', 'pending_approval'], ['pair', 'pending_approval'],
    ['create_user', 'error'], ['pair', 'error'], ['modify_user', 'error'],
  ])('surfaces approval pending at %s for status %s', async (stage, status) => {
    const s = setup(stage === 'modify_user' ? [{ id: 'teams-alice-1j42jy', type: 'user' }] : []);
    if (stage === 'pair') s.callTool.mockResolvedValueOnce({ result: { status: 'success' }, conversationId: 'c' });
    s.callTool.mockResolvedValueOnce({ result: { status, error: 'Policy requires_approval' }, conversationId: 'c' });
    const send = vi.fn(async (_value: unknown) => undefined);
    await handleInbound(
      { ...s, stateStore: { save: s.save } as never,
        options: resolveOptions({ stateDir: '/tmp/x', unknownUserPolicy: 'auto-provision' }) },
      { activity: activity('hello', 'aad-7'), send },
    );
    expect(send.mock.calls[0][0]).toHaveProperty('text', expect.stringMatching(/approval.*pending/i));
    expect(s.streamTool).not.toHaveBeenCalled();
    expect(s.callTool).toHaveBeenCalledTimes(stage === 'pair' ? 2 : 1);
  });

  it('retries pairing without recreating an already provisioned user', async () => {
    const s = setup([{ id: 'teams-alice-1j42jy', name: 'Alice', type: 'user', tools: { search: 'auto' }, status: 'active' }]);
    const send = vi.fn(async (_value: unknown) => undefined);
    await handleInbound(
      { ...s, stateStore: { save: s.save } as never,
        options: resolveOptions({ stateDir: '/tmp/x', unknownUserPolicy: 'auto-provision' }) },
      { activity: activity('hello', 'aad-7'), send },
    );
    expect(s.callTool).toHaveBeenCalledTimes(2);
    expect(s.callTool).toHaveBeenNthCalledWith(1, 'operator', 'modify_user', {
      id: 'teams-alice-1j42jy', tools: { search: 'auto', communicate: 'auto', approval_response: 'auto' },
    });
    expect(s.callTool).toHaveBeenNthCalledWith(2, 'operator', 'pair', expect.objectContaining({
      participantId: 'teams-alice-1j42jy', externalId: 'aad-7',
    }));
    expect(send.mock.calls[0][0]).toHaveProperty('text', expect.stringMatching(/Welcome/));
  });

  it.each([
    { type: 'user', identities: [{ connector: 'teams', externalId: 'someone-else' }] },
    { type: 'agent', identities: [] },
  ])('rejects an existing conflicting participant: $type', async (participant) => {
    const s = setup([{ id: 'teams-alice-1j42jy', name: 'Alice', status: 'active', ...participant }]);
    const send = vi.fn(async (_value: unknown) => undefined);
    await handleInbound(
      { ...s, stateStore: { save: s.save } as never,
        options: resolveOptions({ stateDir: '/tmp/x', unknownUserPolicy: 'auto-provision' }) },
      { activity: activity('hello', 'aad-7'), send },
    );
    expect(s.callTool).not.toHaveBeenCalled();
    expect(send).toHaveBeenCalledWith(
      expect.objectContaining({ text: 'Participant teams-alice-1j42jy is already in use; ask an administrator to pair you.' }),
    );
    expect(s.ctx.registry.setActive).not.toHaveBeenCalled();
  });

  it.each(['create_user', 'pair'])('reports a rejected %s call without resolving the sender', async (stage) => {
    const s = setup();
    if (stage === 'pair') s.callTool.mockResolvedValueOnce({ result: { status: 'success' }, conversationId: 'c' });
    s.callTool.mockRejectedValueOnce(new Error('tool unavailable'));
    const send = vi.fn(async (_value: unknown) => undefined);
    await handleInbound(
      { ...s, stateStore: { save: s.save } as never,
        options: resolveOptions({ stateDir: '/tmp/x', unknownUserPolicy: 'auto-provision' }) },
      { activity: activity('hello'), send },
    );
    expect(send.mock.calls[0][0]).toHaveProperty('text', `${stage === 'pair' ? 'Pairing' : 'Provisioning'} failed: tool unavailable`);
    expect(s.streamTool).not.toHaveBeenCalled();
    expect(s.ctx.registry.setActive).not.toHaveBeenCalled();
  });
});
