import { describe, expect, it, vi } from 'vitest';
import type { ToolContext, ToolResult } from '@legion-collective/core';
import type { BotPool, TeamsAppLike } from './bot-pool.js';
import type { CollectiveLike } from './directory.js';
import { resolveOptions } from './options.js';
import { emptyState, type StateStore, type TeamsState } from './state.js';
import { buildTeamsTools, sendCardTool, postReactionTool, listConversationsTool } from './tools.js';
import { recordMessage } from './tool-state.js';

function setup(
  args: {
    participants?: Array<Record<string, unknown>>;
    conversations?: TeamsState['conversations'];
    options?: Record<string, unknown>;
  } = {},
) {
  const participants = args.participants ?? [
    { id: 'assistant', identities: [{ connector: 'teams-bot', externalId: 'app' }] },
    { id: 'alice', identities: [{ connector: 'teams', externalId: 'aad-1' }] },
    { id: 'team-eng', identities: [{ connector: 'teams-group', externalId: '19:g@thread.v2' }] },
  ];
  const sent: Array<{ conversationId: string; activity: unknown }> = [];
  const reactions = {
    add: vi.fn(async () => undefined),
    delete: vi.fn(async () => undefined),
  };
  const create = vi.fn(async () => ({ id: 'conv-new' }));
  const app: TeamsAppLike = {
    initialize: async () => undefined,
    start: async () => undefined,
    stop: async () => undefined,
    send: async (conversationId, activity) => {
      sent.push({ conversationId, activity });
      return { id: `act-${sent.length}` };
    },
    api: {
      conversations: {
        create,
        updateActivity: async () => undefined,
      },
      reactions,
    },
    on: () => undefined,
  };
  const collective = {
    list: () => participants,
    get: (id: string) => participants.find((p) => p.id === id),
  } as unknown as CollectiveLike;
  const pool = {
    getByAppId: (appId: string) => (appId === 'app' ? app : undefined),
    ownerOf: (appId: string) => (appId === 'app' ? 'assistant' : undefined),
  } as unknown as BotPool;
  const state: TeamsState = { ...emptyState(), conversations: args.conversations ?? {} };
  const deps = {
    collective,
    pool,
    state,
    store: { save: vi.fn(async () => undefined) } as unknown as StateStore,
    options: resolveOptions({ stateDir: '/tmp/x', tenantId: 'tenant-1', ...args.options }),
  };
  const context = { participant: { id: 'assistant' } } as unknown as ToolContext;
  return { deps, context, app, sent, reactions, create };
}

async function run(tool: ReturnType<typeof sendCardTool>, context: ToolContext, args: unknown) {
  return (await tool.execute(args, context)) as ToolResult;
}

describe('teams__send_card', () => {
  it('sends an Adaptive Card to a participant and records the activity id', async () => {
    const s = setup({
      conversations: {
        'app:alice': {
          conversationId: 'conv-1',
          conversationType: 'personal',
          updatedAt: '2026-01-01T00:00:00.000Z',
        },
      },
    });

    const result = await run(sendCardTool(s.deps), s.context, {
      participantId: 'alice',
      card: { type: 'AdaptiveCard', version: '1.4', body: [] },
    });

    expect(result.status).toBe('success');
    expect(s.sent[0]?.conversationId).toBe('conv-1');
    const activity = s.sent[0]?.activity as {
      attachments?: Array<{ contentType: string }>;
    };
    expect(activity.attachments?.[0]?.contentType).toBe('application/vnd.microsoft.card.adaptive');
    expect(s.deps.state.messages['app:act-1']).toMatchObject({
      direction: 'outbound',
      conversationId: 'conv-1',
    });
    expect(result.data).toMatchObject({ conversationId: 'conv-1', activityId: 'act-1' });
  });

  it('sends to a raw Teams conversation that belongs to the caller', async () => {
    const s = setup({
      conversations: {
        'app:alice': {
          conversationId: 'conv-1',
          conversationType: 'personal',
          updatedAt: '2026-01-01T00:00:00.000Z',
        },
      },
    });
    const result = await run(sendCardTool(s.deps), s.context, {
      conversationId: 'conv-1',
      card: { type: 'AdaptiveCard', version: '1.4', body: [] },
    });
    expect(result.status).toBe('success');
    expect(s.sent[0]?.conversationId).toBe('conv-1');
  });

  it('rejects a caller without a live bot', async () => {
    const s = setup({
      participants: [{ id: 'alice', identities: [{ connector: 'teams', externalId: 'aad-1' }] }],
    });
    const result = await run(sendCardTool(s.deps), s.context, {
      participantId: 'alice',
      card: { type: 'AdaptiveCard', version: '1.4', body: [] },
    });
    expect(result.status).toBe('error');
    expect(result.error).toContain('no Teams bot');
  });

  it('rejects an unknown recipient', async () => {
    const s = setup();
    const result = await run(sendCardTool(s.deps), s.context, {
      participantId: 'ghost',
      card: { type: 'AdaptiveCard', version: '1.4', body: [] },
    });
    expect(result.status).toBe('error');
    expect(result.error).toContain('not found');
  });

  it('rejects a card larger than maxCardBytes', async () => {
    const s = setup({ options: { maxCardBytes: 10 } });
    const result = await run(sendCardTool(s.deps), s.context, {
      participantId: 'alice',
      card: { type: 'AdaptiveCard', version: '1.4', body: [] },
    });
    expect(result.status).toBe('error');
    expect(result.error).toContain('max');
  });

  it('reports SDK send failures as tool errors', async () => {
    const s = setup({
      conversations: {
        'app:alice': {
          conversationId: 'conv-1',
          conversationType: 'personal',
          updatedAt: '2026-01-01T00:00:00.000Z',
        },
      },
    });
    s.app.send = vi.fn(async () => {
      throw new Error('boom');
    });
    const result = await run(sendCardTool(s.deps), s.context, {
      participantId: 'alice',
      card: { type: 'AdaptiveCard', version: '1.4', body: [] },
    });
    expect(result.status).toBe('error');
    expect(result.error).toContain('boom');
  });

  it('requires exactly one target', async () => {
    const s = setup();
    const result = await run(sendCardTool(s.deps), s.context, {
      card: { type: 'AdaptiveCard', version: '1.4', body: [] },
    });
    expect(result.status).toBe('error');
    expect(result.error).toContain('exactly one');
  });

  it('appends tool buttons to the card actions', async () => {
    const s = setup({
      conversations: {
        'app:alice': {
          conversationId: 'conv-1',
          conversationType: 'personal',
          updatedAt: '2026-01-01T00:00:00.000Z',
        },
      },
    });
    const result = await run(sendCardTool(s.deps), s.context, {
      participantId: 'alice',
      card: { type: 'AdaptiveCard', version: '1.4', body: [] },
      buttons: [
        { title: 'Add 👍', tool: 'teams__post_reaction', args: { reaction: 'like' } },
      ],
    });
    expect(result.status).toBe('success');
    const activity = s.sent[0]?.activity as {
      attachments?: Array<{ content: { actions?: unknown[] } }>;
    };
    expect(activity.attachments?.[0]?.content?.actions).toEqual([
      {
        type: 'Action.Execute',
        title: 'Add 👍',
        data: { action: 'tool', tool: 'teams__post_reaction', args: { reaction: 'like' } },
      },
    ]);
  });

  it('rejects more than six buttons', async () => {
    const s = setup({
      conversations: {
        'app:alice': {
          conversationId: 'conv-1',
          conversationType: 'personal',
          updatedAt: '2026-01-01T00:00:00.000Z',
        },
      },
    });
    const result = await run(sendCardTool(s.deps), s.context, {
      participantId: 'alice',
      card: { type: 'AdaptiveCard', version: '1.4', body: [] },
      buttons: Array.from({ length: 7 }, (_, index) => ({ title: `B${index}`, tool: 't' })),
    });
    expect(result.status).toBe('error');
    expect(result.error).toContain('at most 6');
  });

  it('rejects a button without a title or tool', async () => {
    const s = setup({
      conversations: {
        'app:alice': {
          conversationId: 'conv-1',
          conversationType: 'personal',
          updatedAt: '2026-01-01T00:00:00.000Z',
        },
      },
    });
    const noTitle = await run(sendCardTool(s.deps), s.context, {
      participantId: 'alice',
      card: { type: 'AdaptiveCard', version: '1.4', body: [] },
      buttons: [{ tool: 'teams__post_reaction' }],
    });
    expect(noTitle.status).toBe('error');
    expect(noTitle.error).toContain('title');
    const noTool = await run(sendCardTool(s.deps), s.context, {
      participantId: 'alice',
      card: { type: 'AdaptiveCard', version: '1.4', body: [] },
      buttons: [{ title: 'Go' }],
    });
    expect(noTool.status).toBe('error');
    expect(noTool.error).toContain('tool');
  });

  it('rejects non-object args and resultCard', async () => {
    const s = setup({
      conversations: {
        'app:alice': {
          conversationId: 'conv-1',
          conversationType: 'personal',
          updatedAt: '2026-01-01T00:00:00.000Z',
        },
      },
    });
    const badArgs = await run(sendCardTool(s.deps), s.context, {
      participantId: 'alice',
      card: { type: 'AdaptiveCard', version: '1.4', body: [] },
      buttons: [{ title: 'Go', tool: 't', args: [] }],
    });
    expect(badArgs.status).toBe('error');
    expect(badArgs.error).toContain('args');
    const badCard = await run(sendCardTool(s.deps), s.context, {
      participantId: 'alice',
      card: { type: 'AdaptiveCard', version: '1.4', body: [] },
      buttons: [{ title: 'Go', tool: 't', resultCard: 'nope' }],
    });
    expect(badCard.status).toBe('error');
    expect(badCard.error).toContain('resultCard');
  });

  it('rejects a card whose existing actions plus buttons exceed six', async () => {
    const s = setup({
      conversations: {
        'app:alice': {
          conversationId: 'conv-1',
          conversationType: 'personal',
          updatedAt: '2026-01-01T00:00:00.000Z',
        },
      },
    });
    const result = await run(sendCardTool(s.deps), s.context, {
      participantId: 'alice',
      card: {
        type: 'AdaptiveCard',
        version: '1.4',
        body: [],
        actions: Array.from({ length: 5 }, () => ({
          type: 'Action.Execute',
          title: 'x',
          data: {},
        })),
      },
      buttons: [
        { title: 'A', tool: 't' },
        { title: 'B', tool: 't' },
      ],
    });
    expect(result.status).toBe('error');
    expect(result.error).toContain('at most 6 actions');
  });

  it('enforces maxCardBytes after adding buttons', async () => {
    const s = setup({
      conversations: {
        'app:alice': {
          conversationId: 'conv-1',
          conversationType: 'personal',
          updatedAt: '2026-01-01T00:00:00.000Z',
        },
      },
      options: { maxCardBytes: 200 },
    });
    const result = await run(sendCardTool(s.deps), s.context, {
      participantId: 'alice',
      card: {
        type: 'AdaptiveCard',
        version: '1.4',
        body: [{ type: 'TextBlock', text: 'x'.repeat(100) }],
      },
      buttons: [{ title: 'Go', tool: 'teams__post_reaction', args: { note: 'y'.repeat(150) } }],
    });
    expect(result.status).toBe('error');
    expect(result.error).toContain('max 200');
  });
});

describe('teams__post_reaction', () => {
  it('adds a reaction by Legion message id', async () => {
    const s = setup({
      conversations: {
        'app:alice': {
          conversationId: 'conv-1',
          conversationType: 'personal',
          updatedAt: '2026-01-01T00:00:00.000Z',
        },
      },
    });
    recordMessage(
      s.deps.state,
      {
        appId: 'app',
        conversationId: 'conv-1',
        activityId: 'act-msg',
        direction: 'inbound',
        legionMessageId: 'msg-1',
      },
      10,
    );

    const result = await run(postReactionTool(s.deps), s.context, {
      messageId: 'msg-1',
      reaction: 'like',
    });

    expect(result.status).toBe('success');
    expect(s.reactions.add).toHaveBeenCalledWith('conv-1', 'act-msg', 'like');
  });

  it('removes a reaction when remove is true', async () => {
    const s = setup({
      conversations: {
        'app:alice': {
          conversationId: 'conv-1',
          conversationType: 'personal',
          updatedAt: '2026-01-01T00:00:00.000Z',
        },
      },
    });
    recordMessage(
      s.deps.state,
      {
        appId: 'app',
        conversationId: 'conv-1',
        activityId: 'act-msg',
        direction: 'inbound',
        legionMessageId: 'msg-1',
      },
      10,
    );

    const result = await run(postReactionTool(s.deps), s.context, {
      messageId: 'msg-1',
      reaction: 'heart',
      remove: true,
    });

    expect(result.status).toBe('success');
    expect(s.reactions.delete).toHaveBeenCalledWith('conv-1', 'act-msg', 'heart');
  });

  it('accepts a raw activity id with a conversation', async () => {
    const s = setup({
      conversations: {
        'app:alice': {
          conversationId: 'conv-1',
          conversationType: 'personal',
          updatedAt: '2026-01-01T00:00:00.000Z',
        },
      },
    });

    const result = await run(postReactionTool(s.deps), s.context, {
      messageId: 'act-raw',
      reaction: 'launch',
      conversationId: 'conv-1',
    });

    expect(result.status).toBe('success');
    expect(s.reactions.add).toHaveBeenCalledWith('conv-1', 'act-raw', 'launch');
  });

  it('rejects an unknown message', async () => {
    const s = setup();
    const result = await run(postReactionTool(s.deps), s.context, {
      messageId: 'nope',
      reaction: 'like',
    });
    expect(result.status).toBe('error');
    expect(result.error).toContain('not found');
  });

  it('rejects a message recorded for another bot', async () => {
    const s = setup();
    recordMessage(
      s.deps.state,
      {
        appId: 'other',
        conversationId: 'conv-1',
        activityId: 'act-msg',
        direction: 'inbound',
        legionMessageId: 'msg-1',
      },
      10,
    );
    const result = await run(postReactionTool(s.deps), s.context, {
      messageId: 'msg-1',
      reaction: 'like',
    });
    expect(result.status).toBe('error');
    expect(result.error).toContain('another bot');
  });

  it('reports SDK reaction failures as tool errors', async () => {
    const s = setup({
      conversations: {
        'app:alice': {
          conversationId: 'conv-1',
          conversationType: 'personal',
          updatedAt: '2026-01-01T00:00:00.000Z',
        },
      },
    });
    recordMessage(
      s.deps.state,
      {
        appId: 'app',
        conversationId: 'conv-1',
        activityId: 'act-msg',
        direction: 'inbound',
        legionMessageId: 'msg-1',
      },
      10,
    );
    s.reactions.add.mockRejectedValue(new Error('reaction failed'));

    const result = await run(postReactionTool(s.deps), s.context, {
      messageId: 'msg-1',
      reaction: 'like',
    });
    expect(result.status).toBe('error');
    expect(result.error).toContain('reaction failed');
  });
});

describe('teams__list_conversations', () => {
  const conversations: TeamsState['conversations'] = {
    'app:alice': {
      conversationId: 'conv-1',
      conversationType: 'personal',
      externalId: 'aad-1',
      displayName: 'Alice',
      updatedAt: '2026-01-02T00:00:00.000Z',
    },
    'app:team-eng': {
      conversationId: '19:g@thread.v2',
      conversationType: 'groupChat',
      externalId: '19:g@thread.v2',
      name: 'Engineering',
      updatedAt: '2026-01-01T00:00:00.000Z',
    },
    'other:alice': {
      conversationId: 'conv-3',
      conversationType: 'personal',
      updatedAt: '2026-01-03T00:00:00.000Z',
    },
  };

  it("lists only the caller bot's conversations, newest first", async () => {
    const s = setup({ conversations });
    const result = await run(listConversationsTool(s.deps), s.context, {});
    expect(result.status).toBe('success');
    const listed = (result.data as { conversations: Array<{ key: string }> }).conversations;
    expect(listed.map((entry) => entry.key)).toEqual(['app:alice', 'app:team-eng']);
  });

  it('filters by conversation type', async () => {
    const s = setup({ conversations });
    const result = await run(listConversationsTool(s.deps), s.context, {
      conversationType: 'groupChat',
    });
    const listed = (result.data as { conversations: Array<{ key: string }> }).conversations;
    expect(listed.map((entry) => entry.key)).toEqual(['app:team-eng']);
  });

  it('caps the limit at 200 and applies a smaller limit', async () => {
    const s = setup({ conversations });
    const limited = await run(listConversationsTool(s.deps), s.context, { limit: 1 });
    expect((limited.data as { conversations: unknown[] }).conversations).toHaveLength(1);

    const capped = await run(listConversationsTool(s.deps), s.context, { limit: 500 });
    expect((capped.data as { conversations: unknown[] }).conversations).toHaveLength(2);
  });

  it('returns an empty list when the bot has no conversations', async () => {
    const s = setup();
    const result = await run(listConversationsTool(s.deps), s.context, {});
    expect(result.data).toEqual({ conversations: [] });
  });

  it('rejects an invalid conversation type', async () => {
    const s = setup();
    const result = await run(listConversationsTool(s.deps), s.context, {
      conversationType: 'channel',
    });
    expect(result.status).toBe('error');
    expect(result.error).toContain('conversationType');
  });
});

describe('buildTeamsTools selection', () => {
  it('returns all three tools by default', () => {
    const s = setup();
    expect(buildTeamsTools(s.deps).map((tool) => tool.name)).toEqual([
      'teams__send_card',
      'teams__post_reaction',
      'teams__list_conversations',
    ]);
  });

  it('returns only the enabled tools', () => {
    const s = setup({ options: { tools: ['reactions'] } });
    expect(buildTeamsTools(s.deps).map((tool) => tool.name)).toEqual(['teams__post_reaction']);
  });

  it('returns no tools when disabled', () => {
    const s = setup({ options: { tools: false } });
    expect(buildTeamsTools(s.deps)).toEqual([]);
  });
});
