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
  it('clears the Legion conversation on /new', async () => {
    const { deps: d, sent } = deps('/new');
    d.state.legionConversations['app:alice'] = 'legacy';
    await expect(handleCommand(d)).resolves.toBe(true);
    expect(d.state.legionConversations['app:alice']).toBeUndefined();
    expect(sent[0]).toMatch(/new conversation/i);
  });

  it('pairs with a valid code via the core pair tool', async () => {
    const { deps: d, callTool, sent } = deps('/pair ABCDEF bob');
    rememberPendingPair(d.state, 'app', 'aad-1', 600);
    const code = Object.keys(d.state.pendingPairs)[0];
    await handleCommand({ ...d, text: `/pair ${code} bob` });
    expect(callTool).toHaveBeenCalledWith('alice', 'pair', {
      participantId: 'bob',
      connector: 'teams',
      externalId: 'aad-1',
    });
    expect(sent[0]).toMatch(/Paired bob/);
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
