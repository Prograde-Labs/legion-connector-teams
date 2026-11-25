import { describe, expect, it, vi } from 'vitest';
import type { ToolContext, ToolResult } from '@legion-collective/core';
import type { BotPool, TeamsAppLike } from './bot-pool.js';
import type { CollectiveLike } from './directory.js';
import { resolveOptions } from './options.js';
import { emptyState, type StateStore, type TeamsState } from './state.js';
import { sendCardTool } from './tools.js';

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
    api: { conversations: { create }, reactions },
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
});
