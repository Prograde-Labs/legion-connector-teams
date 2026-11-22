// src/commands.test.ts
import { describe, expect, it, vi } from 'vitest';
import { handleCommand } from './commands.js';
import { emptyState } from './state.js';
import { resolveOptions } from './options.js';
import { rememberPendingPair } from './pairing.js';
import type { ConnectorContext } from '@legion-collective/core';

function deps(text: string, overrides: Record<string, unknown> = {}) {
  const sent: string[] = [];
  const callTool = vi.fn(async () => ({ result: { status: 'success' }, conversationId: 'c' }));
  return {
    sent,
    callTool,
    deps: {
      appId: 'app',
      participantId: 'alice',
      aadObjectId: 'aad-1',
      text,
      state: emptyState(),
      ctx: { callTool } as unknown as ConnectorContext,
      options: resolveOptions({ stateDir: '/tmp/x' }),
      send: async (value: string) => { sent.push(value); },
      ...overrides,
    },
  };
}

describe('commands', () => {
  it('allows only one in-flight pairing for a code', async () => {
    const { deps: d, sent } = deps('/pair');
    const code = rememberPendingPair(d.state, 'app', 'aad-1', 600);
    let finish!: () => void;
    const wait = new Promise<void>((resolve) => { finish = resolve; });
    const callTool = vi.fn(async () => {
      await wait;
      return { result: { status: 'success' }, conversationId: 'c' };
    });
    const ctx = { callTool } as unknown as ConnectorContext;
    const first = handleCommand({ ...d, ctx, text: `/pair ${code} bob` });
    const second = handleCommand({ ...d, ctx, text: `/pair ${code.toLowerCase()} charlie` });
    try {
      expect(callTool).toHaveBeenCalledOnce();
      expect(sent).toContain('Pairing already in progress.');
    } finally {
      finish();
      await Promise.all([first, second]);
    }
  });

  it('releases the in-flight code after a tool exception', async () => {
    const { deps: d, callTool } = deps('/pair');
    const code = rememberPendingPair(d.state, 'app', 'aad-1', 600);
    callTool.mockImplementationOnce(async () => { throw new Error('tool unavailable'); });
    const request = { ...d, text: `/pair ${code} bob` };
    await expect(handleCommand(request)).rejects.toThrow('tool unavailable');
    expect(d.state.pendingPairs[code]).toBeDefined();
    await expect(handleCommand(request)).resolves.toBe(true);
    expect(callTool).toHaveBeenCalledTimes(2);
    expect(d.state.pendingPairs[code]).toBeUndefined();
  });

  it('clears the Legion conversation on /new', async () => {
    const { deps: d, sent } = deps('/new');
    d.state.legionConversations['app:alice'] = 'legacy';
    await expect(handleCommand(d)).resolves.toBe(true);
    expect(d.state.legionConversations['app:alice']).toBeUndefined();
    expect(sent[0]).toMatch(/new conversation/i);
  });

  it('pairs with a valid code via the core pair tool', async () => {
    const { deps: d, callTool, sent } = deps('/pair ABCDEF bob');
    rememberPendingPair(d.state, 'app', 'aad-9', 600);
    const code = Object.keys(d.state.pendingPairs)[0];
    await handleCommand({ ...d, text: `/pair ${code} bob` });
    expect(callTool).toHaveBeenCalledWith('alice', 'pair', {
      participantId: 'bob',
      connector: 'teams',
      externalId: 'aad-9',
    });
    expect(sent[0]).toMatch(/Paired bob/);
    expect(d.state.pendingPairs[code]).toBeUndefined();
  });

  it.each([
    ['error', 'not authorized', /not authorized/],
    ['pending_approval', undefined, /awaiting approval/],
    ['error', 'Authorization: requires_approval', /awaiting approval/],
  ])('retains the code on %s', async (status, error, message) => {
    const { deps: d, sent } = deps('/pair');
    const code = rememberPendingPair(d.state, 'app', 'aad-1', 600);
    const callTool = vi.fn(async () => ({ result: { status, error }, conversationId: 'c' }));
    await handleCommand({ ...d, text: `/pair ${code.toLowerCase()} bob`, ctx: { callTool } as unknown as ConnectorContext });
    expect(callTool).toHaveBeenCalledOnce();
    expect(d.state.pendingPairs[code]).toBeDefined();
    expect(sent[0]).toMatch(message);
  });

  it.each(['2000-01-01T00:00:00.000Z', 'invalid'])('removes an invalid expiry %s without calling the tool', async (expiresAt) => {
    const { deps: d, callTool } = deps('/pair ABC234 bob');
    d.state.pendingPairs['ABC234'] = { appId: 'app', aadObjectId: 'aad-1', expiresAt };
    await handleCommand(d);
    expect(callTool).not.toHaveBeenCalled();
    expect(d.state.pendingPairs['ABC234']).toBeUndefined();
  });

  it('does not consume a code belonging to another bot', async () => {
    const { deps: d, callTool } = deps('/pair');
    const code = rememberPendingPair(d.state, 'other-app', 'aad-1', 600);
    await handleCommand({ ...d, text: `/pair ${code} bob` });
    expect(callTool).not.toHaveBeenCalled();
    expect(d.state.pendingPairs[code]).toBeDefined();
  });

  it('uses pairingParticipantId when configured', async () => {
    const { deps: d, callTool } = deps('/pair', {
      options: resolveOptions({ stateDir: '/tmp/x', pairingParticipantId: 'pairing-agent' }),
    });
    rememberPendingPair(d.state, 'app', 'aad-1', 600);
    const code = Object.keys(d.state.pendingPairs)[0];
    await handleCommand({ ...d, text: `/pair ${code} bob` });
    expect(callTool).toHaveBeenCalledWith('pairing-agent', 'pair', expect.anything());
  });

  it('reports an invalid code', async () => {
    const { deps: d, sent, callTool } = deps('/pair NOPE bob');
    await handleCommand(d);
    expect(callTool).not.toHaveBeenCalled();
    expect(sent[0]).toMatch(/invalid or expired/);
  });

  it('returns false for plain text', async () => {
    await expect(handleCommand(deps('hello').deps)).resolves.toBe(false);
  });
});
