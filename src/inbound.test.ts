// src/inbound.test.ts
import { describe, expect, it, vi } from 'vitest';
import { handleInbound } from './inbound.js';
import { emptyState } from './state.js';
import { resolveOptions } from './options.js';
import type { BotEntry, CollectiveLike } from './directory.js';
import type { ConnectorContext } from '@legion-collective/core';

function setup(participants: Array<Record<string, unknown>> = []) {
  const state = emptyState();
  const save = vi.fn(async () => undefined);
  const callTool = vi.fn(async (): Promise<{ result: { status: string; error?: string }; conversationId: string }> =>
    ({ result: { status: 'success' }, conversationId: 'c' }));
  const streamTool = vi.fn(async () => ({
    conversationId: 'conv-1',
    gen: (async function* () {
      yield { type: 'text_delta', delta: 'Hi' };
      yield { type: 'stream:done', result: { status: 'success', data: {} } };
    })(),
  }));
  const collective = {
    list: () => participants,
    get: (id: string) => participants.find((p) => p.id === id),
    findByIdentity: (connector: string, externalId: string) =>
      participants.find((p) =>
        (p.identities as Array<{ connector: string; externalId: string }> | undefined)?.some(
          (i) => i.connector === connector && i.externalId === externalId,
        ),
      ),
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
  conversation: { id: 'conv-1' },
  channelData: { tenant: { id: 'tenant-1' } },
  serviceUrl: 'https://smba.trafficmanager.net/teams/',
  text,
});

describe('handleInbound', () => {
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
    expect(s.state.legionConversations['app:alice']).toBe('conv-1');
    expect(s.state.conversations['app:alice'].conversationId).toBe('conv-1');
    expect(s.save).toHaveBeenCalled();
  });

  it('issues a pairing code for an unknown user', async () => {
    const s = setup();
    const send = vi.fn(async (_value: unknown) => undefined);
    await handleInbound(
      { bot: s.bot, ctx: s.ctx, collective: s.collective, state: s.state,
        stateStore: { save: s.save } as never, options: resolveOptions({ stateDir: '/tmp/x' }) },
      { activity: activity('hello', 'aad-9'), send },
    );
    expect(send.mock.calls[0][0]).toMatch(/Not registered/);
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
    }));
    expect(s.callTool).toHaveBeenCalledWith('operator', 'pair', expect.objectContaining({
      connector: 'teams', externalId: 'aad-7',
    }));
    expect(send.mock.calls[0][0]).toMatch(/Welcome/);
  });

  it('tells rejected users how to obtain access', async () => {
    const s = setup();
    const send = vi.fn(async (_value: unknown) => undefined);
    await handleInbound(
      { ...s, stateStore: { save: s.save } as never,
        options: resolveOptions({ stateDir: '/tmp/x', unknownUserPolicy: 'reject' }) },
      { activity: activity('hello'), send },
    );
    expect(send.mock.calls[0][0]).toMatch(/administrator.*register/i);
    expect(send.mock.calls[0][0]).toMatch(/aad-1/);
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

  it.each(['create_user', 'pair'])('surfaces pending approval for %s despite an error', async (stage) => {
    const s = setup();
    if (stage === 'pair') s.callTool.mockResolvedValueOnce({ result: { status: 'success' }, conversationId: 'c' });
    s.callTool.mockResolvedValueOnce({ result: { status: 'pending_approval', error: 'generic error' }, conversationId: 'c' });
    const send = vi.fn(async (_value: unknown) => undefined);
    await handleInbound(
      { ...s, stateStore: { save: s.save } as never,
        options: resolveOptions({ stateDir: '/tmp/x', unknownUserPolicy: 'auto-provision' }) },
      { activity: activity('hello'), send },
    );
    expect(send.mock.calls[0][0]).toMatch(/approval.*pending/i);
    expect(s.streamTool).not.toHaveBeenCalled();
    expect(s.callTool).toHaveBeenCalledTimes(stage === 'pair' ? 2 : 1);
  });

  it('retries pairing without recreating an already provisioned user', async () => {
    const s = setup([{ id: 'teams-alice-1j42jy', name: 'Alice', type: 'user', tools: {}, status: 'active' }]);
    const send = vi.fn(async (_value: unknown) => undefined);
    await handleInbound(
      { ...s, stateStore: { save: s.save } as never,
        options: resolveOptions({ stateDir: '/tmp/x', unknownUserPolicy: 'auto-provision' }) },
      { activity: activity('hello', 'aad-7'), send },
    );
    expect(s.callTool).toHaveBeenCalledTimes(1);
    expect(s.callTool).toHaveBeenCalledWith('operator', 'pair', expect.objectContaining({
      participantId: 'teams-alice-1j42jy', externalId: 'aad-7',
    }));
    expect(send.mock.calls[0][0]).toMatch(/Welcome/);
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
    expect(send.mock.calls[0][0]).toBe(`${stage === 'pair' ? 'Pairing' : 'Provisioning'} failed: tool unavailable`);
    expect(s.streamTool).not.toHaveBeenCalled();
    expect(s.ctx.registry.setActive).not.toHaveBeenCalled();
  });
});
