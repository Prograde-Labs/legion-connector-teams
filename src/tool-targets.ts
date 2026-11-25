import type { BotPool, TeamsAppLike } from './bot-pool.js';
import type { CollectiveLike } from './directory.js';
import { describeError } from './errors.js';
import type { ResolvedOptions } from './options.js';
import type { StateStore, TeamsState } from './state.js';
import { lookupMessageByActivityId, lookupMessageByLegionId } from './tool-state.js';

export interface TeamsToolDeps {
  collective: CollectiveLike;
  pool: BotPool;
  state: TeamsState;
  store: StateStore;
  options: ResolvedOptions;
}

export type Resolved<T> = { ok: true; value: T } | { ok: false; error: string };

export interface CallerBot {
  appId: string;
  app: TeamsAppLike;
}

function fail<T>(error: string): Resolved<T> {
  return { ok: false, error };
}

function succeed<T>(value: T): Resolved<T> {
  return { ok: true, value };
}

export function resolveCallerBot(deps: TeamsToolDeps, participantId: string): Resolved<CallerBot> {
  const participant = deps.collective.get(participantId);
  const identity = participant?.identities?.find((item) => item.connector === 'teams-bot');
  if (!identity) {
    return fail(`caller '${participantId}' has no Teams bot identity`);
  }

  const appId = identity.externalId;
  if (deps.pool.ownerOf(appId) !== participantId) {
    return fail(`caller '${participantId}' has no live Teams bot`);
  }
  const app = deps.pool.getByAppId(appId);
  if (!app) {
    return fail(`Teams app '${appId}' is not running`);
  }
  return succeed({ appId, app });
}

function conversationForApp(
  state: TeamsState,
  appId: string,
  conversationId: string,
): string | undefined {
  for (const [key, record] of Object.entries(state.conversations)) {
    if (key.startsWith(`${appId}:`) && record.conversationId === conversationId) return key;
  }
  return undefined;
}

export async function resolveConversationTarget(
  deps: TeamsToolDeps,
  appId: string,
  app: TeamsAppLike,
  args: { participantId?: string; conversationId?: string },
): Promise<Resolved<{ conversationId: string; key?: string }>> {
  if (args.conversationId !== undefined) {
    const key = conversationForApp(deps.state, appId, args.conversationId);
    if (!key) {
      return fail(`conversation '${args.conversationId}' does not belong to the caller's bot`);
    }
    return succeed({ conversationId: args.conversationId, key });
  }

  const participantId = args.participantId;
  if (participantId === undefined) {
    return fail('a participantId or conversationId is required');
  }
  const participant = deps.collective.get(participantId);
  if (!participant) {
    return fail(`participant '${participantId}' not found`);
  }

  const group = participant.identities?.find((item) => item.connector === 'teams-group');
  if (group) {
    const key = `${appId}:${participantId}:${group.externalId}`;
    const record = deps.state.conversations[key];
    return succeed({ conversationId: record?.conversationId ?? group.externalId, key });
  }

  const user = participant.identities?.find((item) => item.connector === 'teams');
  if (!user) {
    return fail(`participant '${participantId}' has no Teams identity`);
  }

  const key = `${appId}:${participantId}`;
  const existing = deps.state.conversations[key];
  if (existing) {
    return succeed({ conversationId: existing.conversationId, key });
  }

  try {
    const conversation = await app.api.conversations.create({
      members: [{ id: user.externalId }],
      tenantId: deps.options.tenantId,
      channelData: { tenant: { id: deps.options.tenantId } },
    });
    deps.state.conversations[key] = {
      conversationId: conversation.id,
      conversationType: 'personal',
      externalId: user.externalId,
      tenantId: deps.options.tenantId,
      updatedAt: new Date().toISOString(),
    };
    return succeed({ conversationId: conversation.id, key });
  } catch (error) {
    return fail(`Teams conversation creation failed: ${describeError(error)}`);
  }
}

export function resolveMessageTarget(
  deps: TeamsToolDeps,
  appId: string,
  args: { messageId: string; conversationId?: string },
): Resolved<{ conversationId: string; activityId: string }> {
  const legionRecord = lookupMessageByLegionId(deps.state, args.messageId);
  if (legionRecord) {
    if (legionRecord.appId !== appId) {
      return fail(`message '${args.messageId}' belongs to another bot`);
    }
    return succeed({
      conversationId: legionRecord.conversationId,
      activityId: legionRecord.activityId,
    });
  }

  const record = lookupMessageByActivityId(deps.state, args.messageId, appId, args.conversationId);
  if (record) {
    return succeed({ conversationId: record.conversationId, activityId: record.activityId });
  }

  if (args.conversationId !== undefined) {
    const key = conversationForApp(deps.state, appId, args.conversationId);
    if (!key) {
      return fail(`conversation '${args.conversationId}' does not belong to the caller's bot`);
    }
    return succeed({ conversationId: args.conversationId, activityId: args.messageId });
  }

  return fail(`message '${args.messageId}' not found`);
}
