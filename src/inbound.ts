// src/inbound.ts
import type { ConnectorContext } from '@legion-collective/core';
import type { CollectiveLike } from './directory.js';
import { MessageActivityInput, TypingActivityInput } from '@microsoft/teams.api';
import { handleCommand } from './commands.js';
import { runCommunicate, type TeamsStreamer } from './communicate.js';
import type { BotEntry } from './directory.js';
import { resolveIdentity } from './identity.js';
import type { ResolvedOptions } from './options.js';
import type { StateStore, TeamsState } from './state.js';

export interface TeamsActivity {
  from?: { id?: string; name?: string; aadObjectId?: string };
  conversation?: { id?: string };
  channelData?: { tenant?: { id?: string } };
  serviceUrl?: string;
  text?: string;
  value?: { action?: { data?: Record<string, unknown> } };
}

export interface TeamsContextLike {
  activity: TeamsActivity;
  send: (activity: MessageActivityInput | TypingActivityInput | string) => Promise<unknown>;
  stream?: TeamsStreamer;
}

export interface InboundDeps {
  bot: BotEntry;
  ctx: ConnectorContext;
  collective: CollectiveLike;
  state: TeamsState;
  stateStore: StateStore;
  options: ResolvedOptions;
}

function slug(value: string): string {
  return value.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '').slice(0, 24) || 'user';
}

function hash6(value: string): string {
  let hash = 0;
  for (const char of value) hash = (hash * 31 + char.charCodeAt(0)) >>> 0;
  return hash.toString(36).padStart(6, '0').slice(0, 6);
}

export function provisionedUserId(name: string, aadObjectId: string): string {
  return `teams-${slug(name)}-${hash6(aadObjectId)}`;
}

function upsertConversation(
  state: TeamsState,
  bot: BotEntry,
  userId: string,
  aadObjectId: string,
  displayName: string | undefined,
  activity: TeamsActivity,
): void {
  state.conversations[`${bot.appId}:${userId}`] = {
    conversationId: activity.conversation?.id ?? '',
    tenantId: activity.channelData?.tenant?.id,
    serviceUrl: activity.serviceUrl,
    aadObjectId,
    displayName,
    updatedAt: new Date().toISOString(),
  };
}

async function autoProvision(
  deps: InboundDeps,
  aadObjectId: string,
  displayName: string | undefined,
  context: TeamsContextLike,
): Promise<void> {
  const name = displayName ?? 'Teams user';
  const id = provisionedUserId(name, aadObjectId);
  const existing = deps.collective.get(id);

  if (existing && (existing.type !== 'user' || existing.identities?.some(
    (identity) => identity.connector === 'teams' && identity.externalId !== aadObjectId,
  ))) {
    await context.send(`Participant ${id} is already in use; ask an administrator to pair you.`);
    return;
  }

  if (!existing) {
    let created;
    try {
      created = await deps.ctx.callTool(deps.options.provisionerId, 'create_user', { id, name });
    } catch (error) {
      await context.send(`Provisioning failed: ${error instanceof Error ? error.message : String(error)}`);
      return;
    }
    if (created.result.status === 'pending_approval') {
      await context.send('Provisioning approval is pending.');
      return;
    }
    if (created.result.status !== 'success') {
      await context.send(`Provisioning failed: ${created.result.error ?? created.result.status}`);
      return;
    }
  }

  let paired;
  try {
    paired = await deps.ctx.callTool(
      deps.options.pairingParticipantId ?? deps.options.provisionerId,
      'pair',
      { participantId: id, connector: 'teams', externalId: aadObjectId },
    );
  } catch (error) {
    await context.send(`Pairing failed: ${error instanceof Error ? error.message : String(error)}`);
    return;
  }
  if (paired.result.status === 'pending_approval') {
    await context.send('Pairing approval is pending.');
    return;
  }
  if (paired.result.status !== 'success') {
    await context.send(`Pairing failed: ${paired.result.error ?? paired.result.status}`);
    return;
  }

  await context.send(`Welcome, ${name}. You are now paired as ${id}.`);
}

export async function handleInbound(
  deps: InboundDeps,
  context: TeamsContextLike,
): Promise<void> {
  const activity = context.activity;
  const aadObjectId = activity.from?.aadObjectId?.trim() || activity.from?.id?.trim();
  if (!aadObjectId) {
    await context.send('Unable to read your Teams identity.');
    return;
  }

  const outcome = resolveIdentity({
    collective: deps.collective,
    state: deps.state,
    appId: deps.bot.appId,
    aadObjectId,
    displayName: activity.from?.name,
    options: deps.options,
  });

  switch (outcome.kind) {
    case 'pairing':
      await context.send(
        `Not registered with Legion yet. Ask someone with the \`pair\` tool to run ` +
        `/pair ${outcome.code} <participant-id>. Your Teams ID: ${outcome.aadObjectId}.`,
      );
      await deps.stateStore.save(deps.state);
      return;

    case 'reject':
      await context.send(
        `Not registered with Legion yet. Your Teams ID: ${outcome.aadObjectId}. Ask an administrator to register you.`,
      );
      return;

    case 'auto-provision':
      await autoProvision(deps, outcome.aadObjectId, outcome.displayName, context);
      await deps.stateStore.save(deps.state);
      return;

    case 'resolved':
    case 'fallback': {
      const user = outcome.user;
      upsertConversation(
        deps.state,
        deps.bot,
        user.participantId,
        user.aadObjectId,
        user.displayName,
        activity,
      );
      await deps.stateStore.save(deps.state);
      deps.ctx.registry.setActive(user.participantId, 'teams');

      const text = activity.text?.trim();
      if (!text) {
        await context.send('Empty message.');
        return;
      }

      if (text.startsWith('/')) {
        const handled = await handleCommand({
          appId: deps.bot.appId,
          participantId: user.participantId,
          aadObjectId: user.aadObjectId,
          text,
          state: deps.state,
          ctx: deps.ctx,
          options: deps.options,
          send: async (value) => { await context.send(value); },
        });
        if (handled) {
          await deps.stateStore.save(deps.state);
          return;
        }
      }

      let typing: NodeJS.Timeout | undefined;
      if (deps.options.typingIntervalMs > 0) {
        const tick = () => context.send(new TypingActivityInput()).catch(() => undefined);
        tick();
        typing = setInterval(tick, deps.options.typingIntervalMs);
      }

      try {
        const { conversationId } = await runCommunicate({
          ctx: deps.ctx,
          userId: user.participantId,
          botOwnerId: deps.bot.participantId,
          text,
          conversationId:
            deps.state.legionConversations[`${deps.bot.appId}:${user.participantId}`],
          options: deps.options,
          stream: context.stream,
          send: (activity) => context.send(activity),
        });
        deps.state.legionConversations[`${deps.bot.appId}:${user.participantId}`] =
          conversationId;
        await deps.stateStore.save(deps.state);
      } finally {
        if (typing) clearInterval(typing);
      }
      return;
    }
  }
}
