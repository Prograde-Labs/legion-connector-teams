// src/commands.test.ts
import { describe, expect, it, vi } from 'vitest';
import { handleCommand } from './commands.js';
import { emptyState } from './state.js';
import { resolveOptions } from './options.js';
import { rememberPendingPair } from './pairing.js';
import type { ConnectorContext } from '@legion-collective/core';
import type { IAdaptiveCard } from '@microsoft/teams.cards';

function deps(text: string, overrides: Record<string, unknown> = {}) {
  const sent: string[] = [];
  const cards: IAdaptiveCard[] = [];
  const callTool = vi.fn(async () => ({ result: { status: 'success' }, conversationId: 'c' }));
  return {
    sent,
    cards,
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
      sendCard: async (card: IAdaptiveCard) => { cards.push(card); },
      ...overrides,
    },
  };
}

describe('commands', () => {
  it('allows only one in-flight pairing for a code', async () => {
    const { deps: d, sent } = deps('/pair');
    const code = rememberPendingPair(d.state, { appId: 'app', connector: 'teams', externalId: 'aad-1' }, 600);
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
    const code = rememberPendingPair(d.state, { appId: 'app', connector: 'teams', externalId: 'aad-1' }, 600);
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
    rememberPendingPair(d.state, { appId: 'app', connector: 'teams', externalId: 'aad-9' }, 600);
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
    const code = rememberPendingPair(d.state, { appId: 'app', connector: 'teams', externalId: 'aad-1' }, 600);
    const callTool = vi.fn(async () => ({ result: { status, error }, conversationId: 'c' }));
    await handleCommand({ ...d, text: `/pair ${code.toLowerCase()} bob`, ctx: { callTool } as unknown as ConnectorContext });
    expect(callTool).toHaveBeenCalledOnce();
    expect(d.state.pendingPairs[code]).toBeDefined();
    expect(sent[0]).toMatch(message);
  });

  it.each(['2000-01-01T00:00:00.000Z', 'invalid'])('removes an invalid expiry %s without calling the tool', async (expiresAt) => {
    const { deps: d, callTool } = deps('/pair ABC234 bob');
    d.state.pendingPairs['ABC234'] = {
      appId: 'app', connector: 'teams', externalId: 'aad-1', expiresAt,
    };
    await handleCommand(d);
    expect(callTool).not.toHaveBeenCalled();
    expect(d.state.pendingPairs['ABC234']).toBeUndefined();
  });

  it('does not consume a code belonging to another bot', async () => {
    const { deps: d, callTool } = deps('/pair');
    const code = rememberPendingPair(d.state, { appId: 'other-app', connector: 'teams', externalId: 'aad-1' }, 600);
    await handleCommand({ ...d, text: `/pair ${code} bob` });
    expect(callTool).not.toHaveBeenCalled();
    expect(d.state.pendingPairs[code]).toBeDefined();
  });

  it('uses pairingParticipantId when configured', async () => {
    const { deps: d, callTool } = deps('/pair', {
      options: resolveOptions({ stateDir: '/tmp/x', pairingParticipantId: 'pairing-agent' }),
    });
    rememberPendingPair(d.state, { appId: 'app', connector: 'teams', externalId: 'aad-1' }, 600);
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

  it('clears the group conversation and its target map entry on /new', async () => {
    const { deps: d, sent } = deps('/new', {
      participantId: 'team-eng',
      isGroup: true,
      conversationKey: 'app:team-eng:19:g@thread.v2',
      conversationId: '19:g@thread.v2',
    });
    d.state.legionConversations['app:team-eng:19:g@thread.v2'] = 'legion-1';
    d.state.legionConversationTargets['legion-1'] = 'app:team-eng:19:g@thread.v2';
    await expect(handleCommand(d)).resolves.toBe(true);
    expect(d.state.legionConversations['app:team-eng:19:g@thread.v2']).toBeUndefined();
    expect(d.state.legionConversationTargets['legion-1']).toBeUndefined();
    expect(sent[0]).toMatch(/new conversation/i);
  });

  it('shows group identity details on /whoami', async () => {
    const { deps: d, sent } = deps('/whoami', {
      participantId: 'team-eng',
      senderName: 'Alice',
      isGroup: true,
      conversationKey: 'app:team-eng:19:g@thread.v2',
      conversationId: '19:g@thread.v2',
    });
    await expect(handleCommand(d)).resolves.toBe(true);
    expect(sent[0]).toContain('Participant: team-eng');
    expect(sent[0]).toContain('Group conversation ID: 19:g@thread.v2');
    expect(sent[0]).toContain('Sender: Alice');
    expect(sent[0]).toContain('Bot app: app');
  });

  it('runs /pair as the typing human in a group and uses the pending group identity', async () => {
    const { deps: d, callTool, sent } = deps('/pair', {
      participantId: 'team-eng',
      senderParticipantId: 'alice',
      isGroup: true,
      conversationKey: 'app:team-eng:19:g@thread.v2',
      conversationId: '19:g@thread.v2',
    });
    const code = rememberPendingPair(
      d.state,
      { appId: 'app', connector: 'teams-group', externalId: '19:g@thread.v2' },
      600,
    );
    await handleCommand({ ...d, text: `/pair ${code} team-eng` });
    expect(callTool).toHaveBeenCalledWith('alice', 'pair', {
      participantId: 'team-eng',
      connector: 'teams-group',
      externalId: '19:g@thread.v2',
    });
    expect(sent[0]).toMatch(/Paired team-eng with Teams group identity 19:g@thread\.v2/);
    expect(d.state.pendingPairs[code]).toBeUndefined();
  });

  it('runs /pair as the registered human in a fallback group', async () => {
    const { deps: d, callTool, sent } = deps('/pair', {
      participantId: 'operator',
      senderParticipantId: 'alice',
      isFallback: true,
      isGroup: true,
      conversationKey: 'app:operator:19:g@thread.v2',
      conversationId: '19:g@thread.v2',
    });
    const code = rememberPendingPair(
      d.state,
      { appId: 'app', connector: 'teams-group', externalId: '19:g@thread.v2' },
      600,
    );
    await handleCommand({ ...d, text: `/pair ${code} team-eng` });
    expect(callTool).toHaveBeenCalledWith('alice', 'pair', {
      participantId: 'team-eng',
      connector: 'teams-group',
      externalId: '19:g@thread.v2',
    });
    expect(sent[0]).toMatch(/Paired team-eng/);
    expect(d.state.pendingPairs[code]).toBeUndefined();
  });

  it('still blocks /pair in a personal fallback chat', async () => {
    const { deps: d, callTool, sent } = deps('/pair', {
      participantId: 'alice',
      senderParticipantId: 'alice',
      isFallback: true,
    });
    const code = rememberPendingPair(
      d.state,
      { appId: 'app', connector: 'teams', externalId: 'aad-9' },
      600,
    );
    await handleCommand({ ...d, text: `/pair ${code} bob` });
    expect(callTool).not.toHaveBeenCalled();
    expect(sent[0]).toBe('Only registered users can redeem pairing codes.');
    expect(d.state.pendingPairs[code]).toBeDefined();
  });

  it('mentions group behavior in /help', async () => {
    const { deps: d, sent } = deps('/help');
    await expect(handleCommand(d)).resolves.toBe(true);
    expect(sent[0]).toContain('In a group chat, /new and /whoami act on the group; /pair runs as you.');
  });

  it('rejects /pair from an unregistered human in a group', async () => {
    const { deps: d, callTool, sent } = deps('/pair', {
      participantId: 'team-eng',
      isGroup: true,
      conversationKey: 'app:team-eng:19:g@thread.v2',
      conversationId: '19:g@thread.v2',
    });
    const code = rememberPendingPair(
      d.state,
      { appId: 'app', connector: 'teams-group', externalId: '19:g@thread.v2' },
      600,
    );
    await handleCommand({ ...d, text: `/pair ${code} team-eng` });
    expect(callTool).not.toHaveBeenCalled();
    expect(sent[0]).toBe('Only registered users can redeem pairing codes.');
    expect(d.state.pendingPairs[code]).toBeDefined();
  });

  it('sends the conversation picker card on /switch', async () => {
    const { deps: d, cards } = deps('/switch');
    const callTool = vi.fn(async () => ({
      result: { status: 'success', data: { conversations: [
        { id: 'legion-1', resolvedTitle: 'First', messageCount: 3,
          updatedAt: '2026-11-24T00:00:00.000Z', status: 'active' },
      ] } },
      conversationId: 'c',
    }));
    await handleCommand({ ...d, ctx: { callTool } as unknown as ConnectorContext });
    expect(callTool).toHaveBeenCalledWith('alice', 'list_conversations', {
      participantId: 'alice',
      status: 'active',
    });
    expect(cards).toHaveLength(1);
    expect(JSON.stringify(cards[0])).toContain('First');
  });

  it('switches directly to a conversation id and cleans the old target', async () => {
    const { deps: d, sent } = deps('/switch legion-2');
    d.state.legionConversations['app:alice'] = 'legion-1';
    d.state.legionConversationTargets['legion-1'] = 'app:alice';
    const callTool = vi.fn(async () => ({
      result: { status: 'success', data: { conversations: [
        { id: 'legion-2', resolvedTitle: 'Second', messageCount: 2,
          updatedAt: '2026-11-25T00:00:00.000Z', status: 'active' },
      ] } },
      conversationId: 'c',
    }));
    await handleCommand({ ...d, ctx: { callTool } as unknown as ConnectorContext });
    expect(d.state.legionConversations['app:alice']).toBe('legion-2');
    expect(d.state.legionConversationTargets['legion-1']).toBeUndefined();
    expect(d.state.legionConversationTargets['legion-2']).toBe('app:alice');
    expect(sent[0]).toBe('Switched to Second (2 messages).');
  });

  it('reports the current conversation on a direct switch', async () => {
    const { deps: d, sent } = deps('/switch legion-1');
    d.state.legionConversations['app:alice'] = 'legion-1';
    const callTool = vi.fn(async () => ({
      result: { status: 'success', data: { conversations: [
        { id: 'legion-1', resolvedTitle: 'First', messageCount: 3,
          updatedAt: '2026-11-24T00:00:00.000Z', status: 'active' },
      ] } },
      conversationId: 'c',
    }));
    await handleCommand({ ...d, ctx: { callTool } as unknown as ConnectorContext });
    expect(sent[0]).toBe('Already on First.');
  });

  it('reports an unknown conversation id', async () => {
    const { deps: d, sent } = deps('/switch missing');
    const callTool = vi.fn(async () => ({
      result: { status: 'success', data: { conversations: [] } },
      conversationId: 'c',
    }));
    await handleCommand({ ...d, ctx: { callTool } as unknown as ConnectorContext });
    expect(sent[0]).toBe('Conversation not found. Use /switch to list active conversations.');
  });

  it('reports a listing failure', async () => {
    const { deps: d, sent } = deps('/switch');
    const callTool = vi.fn(async () => ({
      result: { status: 'error', error: 'not authorized' },
      conversationId: 'c',
    }));
    await handleCommand({ ...d, ctx: { callTool } as unknown as ConnectorContext });
    expect(sent[0]).toBe('Unable to list conversations: not authorized');
  });

  it('uses the group key for /switch in a group chat', async () => {
    const { deps: d, sent } = deps('/switch legion-2', {
      participantId: 'team-eng',
      isGroup: true,
      conversationKey: 'app:team-eng:19:g@thread.v2',
      conversationId: '19:g@thread.v2',
    });
    d.state.legionConversations['app:team-eng:19:g@thread.v2'] = 'legion-1';
    d.state.legionConversationTargets['legion-1'] = 'app:team-eng:19:g@thread.v2';
    const callTool = vi.fn(async () => ({
      result: { status: 'success', data: { conversations: [
        { id: 'legion-2', resolvedTitle: 'Second', messageCount: 2,
          updatedAt: '2026-11-25T00:00:00.000Z', status: 'active' },
      ] } },
      conversationId: 'c',
    }));
    await handleCommand({ ...d, ctx: { callTool } as unknown as ConnectorContext });
    expect(callTool).toHaveBeenCalledWith('team-eng', 'list_conversations', expect.anything());
    expect(d.state.legionConversations['app:team-eng:19:g@thread.v2']).toBe('legion-2');
    expect(d.state.legionConversationTargets['legion-1']).toBeUndefined();
    expect(d.state.legionConversationTargets['legion-2']).toBe('app:team-eng:19:g@thread.v2');
    expect(sent[0]).toContain('Switched to Second');
  });

  it('uses the fallback suffix key for /switch in a personal chat', async () => {
    const { deps: d, sent } = deps('/switch legion-2', {
      participantId: 'operator',
      isFallback: true,
      conversationKey: 'app:operator:aad-9',
    });
    d.state.legionConversations['app:operator:aad-9'] = 'legion-1';
    d.state.legionConversationTargets['legion-1'] = 'app:operator:aad-9';
    const callTool = vi.fn(async () => ({
      result: { status: 'success', data: { conversations: [
        { id: 'legion-2', resolvedTitle: 'Second', messageCount: 2,
          updatedAt: '2026-11-25T00:00:00.000Z', status: 'active' },
      ] } },
      conversationId: 'c',
    }));
    await handleCommand({ ...d, ctx: { callTool } as unknown as ConnectorContext });
    expect(callTool).toHaveBeenCalledWith('operator', 'list_conversations', expect.anything());
    expect(d.state.legionConversations['app:operator:aad-9']).toBe('legion-2');
    expect(d.state.legionConversationTargets['legion-1']).toBeUndefined();
    expect(d.state.legionConversationTargets['legion-2']).toBe('app:operator:aad-9');
    expect(sent[0]).toContain('Switched to Second');
  });
});
