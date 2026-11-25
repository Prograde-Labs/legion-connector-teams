import { describe, expect, it, vi } from 'vitest';
import type { BotPool, TeamsAppLike } from './bot-pool.js';
import type { CollectiveLike } from './directory.js';
import { resolveOptions } from './options.js';
import { emptyState, type StateStore, type TeamsState } from './state.js';
import { recordMessage } from './tool-state.js';
import {
  resolveCallerBot,
  resolveConversationTarget,
  resolveMessageTarget,
} from './tool-targets.js';

function setup(
  args: {
    participants?: Array<Record<string, unknown>>;
    conversations?: TeamsState['conversations'];
  } = {},
) {
  const participants = args.participants ?? [
    { id: 'assistant', identities: [{ connector: 'teams-bot', externalId: 'app' }] },
    { id: 'alice', identities: [{ connector: 'teams', externalId: 'aad-1' }] },
    { id: 'team-eng', identities: [{ connector: 'teams-group', externalId: '19:g@thread.v2' }] },
  ];
  const create = vi.fn(async () => ({ id: 'conv-new' }));
  const app: TeamsAppLike = {
    initialize: async () => undefined,
    start: async () => undefined,
    stop: async () => undefined,
    send: async () => ({ id: 'act-1' }),
    api: {
      conversations: {
        create,
        updateActivity: async () => undefined,
      },
      reactions: { add: async () => undefined, delete: async () => undefined },
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
    options: resolveOptions({ stateDir: '/tmp/x', tenantId: 'tenant-1' }),
  };
  return { deps, app, create };
}

describe('resolveCallerBot', () => {
  it("resolves the caller's live bot", () => {
    const { deps, app } = setup();
    expect(resolveCallerBot(deps, 'assistant')).toEqual({
      ok: true,
      value: { appId: 'app', app },
    });
  });

  it('rejects a caller without a teams-bot identity', () => {
    const { deps } = setup();
    const result = resolveCallerBot(deps, 'alice');
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error).toContain('no Teams bot');
  });

  it('rejects a bot owned by another participant', () => {
    const { deps } = setup();
    deps.pool = {
      getByAppId: () => undefined,
      ownerOf: () => 'someone-else',
    } as unknown as BotPool;
    const result = resolveCallerBot(deps, 'assistant');
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error).toContain('no live Teams bot');
  });
});

describe('resolveConversationTarget', () => {
  it('uses a stored personal conversation', async () => {
    const { deps, app } = setup({
      conversations: {
        'app:alice': {
          conversationId: 'conv-1',
          conversationType: 'personal',
          updatedAt: '2026-01-01T00:00:00.000Z',
        },
      },
    });
    expect(await resolveConversationTarget(deps, 'app', app, { participantId: 'alice' })).toEqual({
      ok: true,
      value: { conversationId: 'conv-1', key: 'app:alice' },
    });
  });

  it('creates a personal conversation when none exists', async () => {
    const { deps, app, create } = setup();
    const result = await resolveConversationTarget(deps, 'app', app, { participantId: 'alice' });
    expect(result.ok).toBe(true);
    expect(create).toHaveBeenCalledWith({
      members: [{ id: 'aad-1' }],
      tenantId: 'tenant-1',
      channelData: { tenant: { id: 'tenant-1' } },
    });
    expect(deps.state.conversations['app:alice']?.conversationId).toBe('conv-new');
  });

  it('uses the group identity conversation id', async () => {
    const { deps, app } = setup();
    expect(
      await resolveConversationTarget(deps, 'app', app, { participantId: 'team-eng' }),
    ).toEqual({
      ok: true,
      value: { conversationId: '19:g@thread.v2', key: 'app:team-eng:19:g@thread.v2' },
    });
  });

  it('stores a group conversation record when none exists', async () => {
    const { deps, app } = setup();
    await resolveConversationTarget(deps, 'app', app, { participantId: 'team-eng' });
    const stored = deps.state.conversations['app:team-eng:19:g@thread.v2'];
    expect(stored).toMatchObject({
      conversationId: '19:g@thread.v2',
      conversationType: 'groupChat',
      externalId: '19:g@thread.v2',
    });
  });

  it('falls back to a personal variant record without identities', async () => {
    const { deps, app, create } = setup({
      participants: [
        { id: 'assistant', identities: [{ connector: 'teams-bot', externalId: 'app' }] },
        { id: 'fallback', identities: [] },
      ],
      conversations: {
        'app:fallback:aad-9': {
          conversationId: 'conv-9',
          conversationType: 'personal',
          updatedAt: '2026-01-01T00:00:00.000Z',
        },
      },
    });
    expect(await resolveConversationTarget(deps, 'app', app, { participantId: 'fallback' })).toEqual(
      { ok: true, value: { conversationId: 'conv-9', key: 'app:fallback:aad-9' } },
    );
    expect(create).not.toHaveBeenCalled();
  });

  it('picks the most recently updated personal variant', async () => {
    const { deps, app } = setup({
      participants: [
        { id: 'assistant', identities: [{ connector: 'teams-bot', externalId: 'app' }] },
        { id: 'fallback', identities: [] },
      ],
      conversations: {
        'app:fallback:aad-1': {
          conversationId: 'conv-old',
          conversationType: 'personal',
          updatedAt: '2026-01-01T00:00:00.000Z',
        },
        'app:fallback:aad-2': {
          conversationId: 'conv-new',
          conversationType: 'personal',
          updatedAt: '2026-02-01T00:00:00.000Z',
        },
        'app:fallback:19:g@thread.v2': {
          conversationId: '19:g@thread.v2',
          conversationType: 'groupChat',
          updatedAt: '2026-03-01T00:00:00.000Z',
        },
      },
    });
    expect(await resolveConversationTarget(deps, 'app', app, { participantId: 'fallback' })).toEqual(
      { ok: true, value: { conversationId: 'conv-new', key: 'app:fallback:aad-2' } },
    );
  });

  it('fails when there is no identity and no variant record', async () => {
    const { deps, app } = setup({
      participants: [
        { id: 'assistant', identities: [{ connector: 'teams-bot', externalId: 'app' }] },
        { id: 'fallback', identities: [] },
      ],
    });
    const result = await resolveConversationTarget(deps, 'app', app, { participantId: 'fallback' });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error).toContain('has no Teams identity');
  });

  it('resolves a raw conversation that belongs to the bot', async () => {
    const { deps, app } = setup({
      conversations: {
        'app:alice': {
          conversationId: 'conv-1',
          conversationType: 'personal',
          updatedAt: '2026-01-01T00:00:00.000Z',
        },
      },
    });
    expect(await resolveConversationTarget(deps, 'app', app, { conversationId: 'conv-1' })).toEqual(
      {
        ok: true,
        value: { conversationId: 'conv-1', key: 'app:alice' },
      },
    );
  });

  it('rejects an unknown raw conversation', async () => {
    const { deps, app } = setup();
    const result = await resolveConversationTarget(deps, 'app', app, { conversationId: 'conv-x' });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error).toContain('does not belong');
  });

  it('rejects an unknown participant', async () => {
    const { deps, app } = setup();
    const result = await resolveConversationTarget(deps, 'app', app, { participantId: 'ghost' });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error).toContain('not found');
  });
});

describe('resolveMessageTarget', () => {
  it('resolves a message by Legion id', () => {
    const { deps } = setup({
      conversations: {
        'app:alice': {
          conversationId: 'conv-1',
          conversationType: 'personal',
          updatedAt: '2026-01-01T00:00:00.000Z',
        },
      },
    });
    recordMessage(
      deps.state,
      {
        appId: 'app',
        conversationId: 'conv-1',
        activityId: 'act-1',
        direction: 'inbound',
        legionMessageId: 'msg-1',
      },
      10,
    );
    expect(resolveMessageTarget(deps, 'app', { messageId: 'msg-1' })).toEqual({
      ok: true,
      value: { conversationId: 'conv-1', activityId: 'act-1' },
    });
  });

  it('resolves a message by recorded activity id', () => {
    const { deps } = setup();
    recordMessage(
      deps.state,
      {
        appId: 'app',
        conversationId: 'conv-1',
        activityId: 'act-1',
        direction: 'outbound',
      },
      10,
    );
    expect(resolveMessageTarget(deps, 'app', { messageId: 'act-1' })).toEqual({
      ok: true,
      value: { conversationId: 'conv-1', activityId: 'act-1' },
    });
  });

  it('accepts a raw activity id with a bot conversation', () => {
    const { deps } = setup({
      conversations: {
        'app:alice': {
          conversationId: 'conv-1',
          conversationType: 'personal',
          updatedAt: '2026-01-01T00:00:00.000Z',
        },
      },
    });
    expect(
      resolveMessageTarget(deps, 'app', { messageId: 'act-raw', conversationId: 'conv-1' }),
    ).toEqual({ ok: true, value: { conversationId: 'conv-1', activityId: 'act-raw' } });
  });

  it('rejects a raw activity id with an unknown conversation', () => {
    const { deps } = setup();
    const result = resolveMessageTarget(deps, 'app', {
      messageId: 'act-raw',
      conversationId: 'conv-x',
    });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error).toContain('does not belong');
  });

  it('rejects a message recorded for another bot', () => {
    const { deps } = setup();
    recordMessage(
      deps.state,
      {
        appId: 'other',
        conversationId: 'conv-1',
        activityId: 'act-1',
        direction: 'inbound',
        legionMessageId: 'msg-1',
      },
      10,
    );
    const result = resolveMessageTarget(deps, 'app', { messageId: 'msg-1' });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error).toContain('another bot');
  });
});
