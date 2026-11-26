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

import type { IAdaptiveCard } from '@microsoft/teams.cards';
import {
  CONVERSATION_PAGE_SIZE,
  conversationCard,
  conversationErrorCard,
  switchedCard,
  type ConversationEntry,
} from './conversations.js';

function entry(id: string, title: string, updatedAt = '2026-11-25T00:00:00.000Z'): ConversationEntry {
  return { id, title, messageCount: 3, updatedAt, status: 'active' };
}

describe('conversationCard', () => {
  it('shows five conversations and a More button', () => {
    const entries = Array.from({ length: 7 }, (_, index) => entry(`c${index}`, `Conversation ${index}`));
    const card = conversationCard(entries, 'c1', 0, 'app:alice') as unknown as {
      body: Array<{ text?: string }>;
      actions: Array<{ title: string; data: Record<string, unknown> }>;
    };
    expect(CONVERSATION_PAGE_SIZE).toBe(5);
    expect(card.actions).toHaveLength(6);
    expect(card.actions.slice(0, 5).map((action) => action.data.conversationId))
      .toEqual(['c0', 'c1', 'c2', 'c3', 'c4']);
    expect(card.actions[5]).toEqual({
      type: 'Action.Execute',
      title: 'More…',
      data: { action: 'switch', offset: 5, key: 'app:alice' },
    });
    expect(card.body.some((block) => block.text === 'Showing 1–5 of 7')).toBe(true);
    expect(JSON.stringify(card)).toContain('· current');
  });

  it('shows the final page without More', () => {
    const entries = Array.from({ length: 7 }, (_, index) => entry(`c${index}`, `Conversation ${index}`));
    const card = conversationCard(entries, undefined, 5, 'app:alice') as unknown as {
      body: Array<{ text?: string }>;
      actions: unknown[];
    };
    expect(card.actions).toHaveLength(2);
    expect(card.body.some((block) => block.text === 'Showing 6–7 of 7')).toBe(true);
  });

  it('marks a current conversation outside the list', () => {
    const card = conversationCard([entry('c0', 'Only')], 'gone', 0, 'app:alice') as unknown as {
      body: Array<{ text?: string }>;
    };
    expect(card.body.some((block) => block.text === 'Current: gone (not in the active list)'))
      .toBe(true);
  });

  it('renders an empty state with no actions', () => {
    const card = conversationCard([], undefined, 0, 'app:alice') as unknown as {
      body: Array<{ text?: string }>;
      actions?: unknown[];
    };
    expect(card.body.some((block) => block.text === '🔀 Switch conversation')).toBe(true);
    expect(card.body.some((block) => block.text === 'No conversations yet. Send a message to start one.'))
      .toBe(true);
    expect(card.actions).toBeUndefined();
  });

  it('truncates long titles and stays under the default card limit', () => {
    const entries = Array.from({ length: 5 }, (_, index) =>
      entry(`c${index}`, `${index}`.padEnd(80, 'x')));
    const card = conversationCard(entries, undefined, 0, 'app:alice');
    const serialized = JSON.stringify(card);
    expect(serialized.length).toBeLessThan(28000);
    expect(serialized).toContain('…');
    expect(serialized).not.toContain('x'.repeat(60));
  });
});

describe('switchedCard', () => {
  it('shows the new binding', () => {
    const card = switchedCard(entry('c1', 'Project X')) as unknown as {
      body: Array<{ text?: string }>;
    };
    expect(card.body.some((block) => block.text === '✅ Switched conversation')).toBe(true);
    expect(JSON.stringify(card)).toContain('Project X');
    expect(JSON.stringify(card)).toContain('Future messages in this chat go to this conversation.');
  });

  it('shows the already-current state', () => {
    const card = switchedCard(entry('c1', 'Project X'), true) as unknown as {
      body: Array<{ text?: string }>;
    };
    expect(card.body.some((block) => block.text === '✅ Already on this conversation')).toBe(true);
  });
});

describe('conversationErrorCard', () => {
  it('shows a warning with no actions', () => {
    const card = conversationErrorCard('Conversation not found.') as unknown as {
      body: Array<{ text?: string }>;
      actions?: unknown[];
    };
    expect(card.body[0]?.text).toBe('⚠️ Conversation not found.');
    expect(card.actions).toBeUndefined();
  });
});

it('keeps the card typed as IAdaptiveCard', () => {
  const card: IAdaptiveCard = conversationCard([], undefined, 0, 'app:alice');
  expect(card.type).toBe('AdaptiveCard');
});
