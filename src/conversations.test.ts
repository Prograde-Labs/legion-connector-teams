// src/conversations.test.ts
import { describe, expect, it, vi } from 'vitest';
import type { ConnectorContext } from '@legion-collective/core';
import {
  ConversationToolError,
  bindConversation,
  conversationErrorText,
  listConversations,
} from './conversations.js';
import { emptyState } from './state.js';

function ctx(result: unknown): ConnectorContext {
  return {
    callTool: vi.fn(async () => ({ result, conversationId: 'c' })),
  } as unknown as ConnectorContext;
}

describe('listConversations', () => {
  it('calls list_conversations as the participant and maps active conversations', async () => {
    const context = ctx({
      status: 'success',
      data: {
        conversations: [
          {
            id: 'older',
            resolvedTitle: 'Older',
            messageCount: 2,
            updatedAt: '2026-11-24T10:00:00.000Z',
            status: 'active',
          },
          {
            id: 'newer',
            sharedTitle: 'Newer shared',
            messageCount: 5,
            updatedAt: '2026-11-25T10:00:00.000Z',
            status: 'active',
          },
        ],
      },
    });
    const entries = await listConversations(context, 'alice');
    expect(context.callTool).toHaveBeenCalledWith('alice', 'list_conversations', {
      participantId: 'alice',
      status: 'active',
    });
    expect(entries.map((entry) => entry.id)).toEqual(['newer', 'older']);
    expect(entries[0]).toMatchObject({
      title: 'Newer shared',
      messageCount: 5,
      updatedAt: '2026-11-25T10:00:00.000Z',
      status: 'active',
    });
  });

  it('falls back through resolved, shared, title, then id', async () => {
    const context = ctx({
      status: 'success',
      data: {
        conversations: [
          { id: 'a', title: 'Plain title', messageCount: 1, updatedAt: '2026-11-25T00:00:00.000Z' },
          { id: 'b', messageCount: 1, updatedAt: '2026-11-25T00:00:00.000Z' },
        ],
      },
    });
    const entries = await listConversations(context, 'alice');
    expect(entries.map((entry) => entry.title)).toEqual(['Plain title', 'b']);
  });

  it('treats a missing conversation list as empty', async () => {
    await expect(listConversations(ctx({ status: 'success', data: {} }), 'alice'))
      .resolves.toEqual([]);
  });

  it('throws a typed error on tool failure', async () => {
    const context = ctx({ status: 'error', error: 'not authorized' });
    const error = await listConversations(context, 'alice').catch((value) => value);
    expect(error).toBeInstanceOf(ConversationToolError);
    expect((error as ConversationToolError).result).toEqual({
      status: 'error',
      error: 'not authorized',
    });
  });
});

describe('conversationErrorText', () => {
  it('uses the plain error for error results', () => {
    expect(conversationErrorText(new ConversationToolError({
      status: 'error',
      error: 'not authorized',
    }))).toBe('not authorized');
  });

  it('uses the approval text for pending results', () => {
    expect(conversationErrorText(new ConversationToolError({ status: 'pending_approval' })))
      .toBe('Awaiting approval.');
  });

  it('uses the rejection message', () => {
    expect(conversationErrorText(new ConversationToolError({
      status: 'rejected',
      message: 'Denied by policy.',
    }))).toBe('Denied by policy.');
  });

  it('falls back to the thrown message', () => {
    expect(conversationErrorText(new Error('network down'))).toBe('network down');
  });
});

describe('bindConversation', () => {
  it('binds a new conversation and cleans the previous target', () => {
    const state = emptyState();
    state.legionConversations['app:alice'] = 'old';
    state.legionConversationTargets['old'] = 'app:alice';

    bindConversation(state, 'app:alice', 'new');

    expect(state.legionConversations['app:alice']).toBe('new');
    expect(state.legionConversationTargets['new']).toBe('app:alice');
    expect(state.legionConversationTargets['old']).toBeUndefined();
  });

  it('does nothing when the conversation is already current', () => {
    const state = emptyState();
    state.legionConversations['app:alice'] = 'same';
    state.legionConversationTargets['same'] = 'app:alice';

    bindConversation(state, 'app:alice', 'same');

    expect(state.legionConversations['app:alice']).toBe('same');
    expect(state.legionConversationTargets['same']).toBe('app:alice');
  });

  it('keeps the previous target when another chat now owns it', () => {
    const state = emptyState();
    state.legionConversations['app:alice'] = 'shared';
    state.legionConversationTargets['shared'] = 'app:bob';

    bindConversation(state, 'app:alice', 'new');

    expect(state.legionConversationTargets['shared']).toBe('app:bob');
  });
});
