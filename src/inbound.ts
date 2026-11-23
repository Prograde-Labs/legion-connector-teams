// src/inbound.ts
import type { ConnectorContext } from '@legion-collective/core';
import type { CollectiveLike } from './directory.js';
import { MessageActivityInput, TypingActivityInput } from '@microsoft/teams.api';
import { handleCommand } from './commands.js';
import { isAwaitingApproval } from './approvals.js';
import { runCommunicate, type TeamsStreamer } from './communicate.js';
import type { BotEntry } from './directory.js';
import { readTeamsId, resolveIdentity, provisionedUserId } from './identity.js';
import type { ResolvedOptions } from './options.js';
import { sendText } from './render.js';
import type { StateStore, TeamsState } from './state.js';

export interface TeamsActivity {
  from?: { id?: string; name?: string; aadObjectId?: string };
  conversation?: { id?: string; conversationType?: string };
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

export function isBotTenant(bot: BotEntry, activity: TeamsActivity): boolean {
  return activity.channelData?.tenant?.id === bot.credentials.tenantId;
}

export const tenantMismatchText = 'This bot is not available in your tenant.';

export interface InboundDeps {
  bot: BotEntry;
  ctx: ConnectorContext;
  collective: CollectiveLike;
  state: TeamsState;
  stateStore: StateStore;
  options: ResolvedOptions;
}

function upsertConversation(
  state: TeamsState,
  key: string,
  aadObjectId: string,
  displayName: string | undefined,
  activity: TeamsActivity,
): void {
  state.conversations[key] = {
    conversationId: activity.conversation?.id ?? '',
    conversationType: 'personal',
    externalId: aadObjectId,
    tenantId: activity.channelData?.tenant?.id,
    serviceUrl: activity.serviceUrl,
    displayName,
    updatedAt: new Date().toISOString(),
  };
}

async function autoProvision(
  deps: InboundDeps,
  aadObjectId: string,
  displayName: string | undefined,
  reply: (text: string) => Promise<void>,
): Promise<void> {
  const name = displayName ?? 'Teams user';
  const id = provisionedUserId(name, aadObjectId);
  const existing = deps.collective.get(id);

  if (existing && (existing.type !== 'user' || existing.status === 'retired' || existing.identities?.some(
    (identity) => identity.connector === 'teams' && identity.externalId !== aadObjectId,
  ))) {
    await reply(`Participant ${id} is already in use; ask an administrator to pair you.`);
    return;
  }

  let provisioned;
  const tools = { ...existing?.tools, communicate: 'auto', approval_response: 'auto' };
  try {
    provisioned = existing
      ? await deps.ctx.callTool(deps.options.provisionerId, 'modify_user', { id, tools })
      : await deps.ctx.callTool(deps.options.provisionerId, 'create_user', { id, name, tools });
  } catch (error) {
    await reply(`Provisioning failed: ${error instanceof Error ? error.message : String(error)}`);
    return;
  }
  if (isAwaitingApproval(provisioned.result)) {
    await reply('Provisioning approval is pending.');
    return;
  }
  if (provisioned.result.status !== 'success') {
    await reply(`Provisioning failed: ${provisioned.result.error ?? provisioned.result.status}`);
    return;
  }

  let paired;
  try {
    paired = await deps.ctx.callTool(
      deps.options.pairingParticipantId ?? deps.options.provisionerId,
      'pair',
      { participantId: id, connector: 'teams', externalId: aadObjectId },
    );
  } catch (error) {
    await reply(`Pairing failed: ${error instanceof Error ? error.message : String(error)}`);
    return;
  }
  if (isAwaitingApproval(paired.result)) {
    await reply('Pairing approval is pending.');
    return;
  }
  if (paired.result.status !== 'success') {
    await reply(`Pairing failed: ${paired.result.error ?? paired.result.status}`);
    return;
  }

  await reply(`Welcome, ${name}. You are now paired as ${id}.`);
}

export async function handleInbound(
  deps: InboundDeps,
  context: TeamsContextLike,
): Promise<void> {
  const activity = context.activity;
  const reply = (text: string) => sendText((message) => context.send(message), deps.options.aiLabel, text);
  if (activity.conversation?.conversationType !== 'personal') {
    await reply('Only personal chats are supported.');
    return;
  }
  if (!isBotTenant(deps.bot, activity)) {
    await reply(tenantMismatchText);
    return;
  }
  const aadObjectId = readTeamsId(activity.from);
  if (!aadObjectId) {
    await reply('Unable to read your Teams identity.');
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
      await deps.stateStore.save(deps.state);
      await reply(
        `Not registered with Legion yet. Ask someone with the \`pair\` tool to run ` +
        `/pair ${outcome.code} <participant-id>. Your Teams ID: ${outcome.aadObjectId}.`,
      );
      return;

    case 'reject':
      await reply(
        `Not registered with Legion yet. Your Teams ID: ${outcome.aadObjectId}. Ask an administrator to register you.`,
      );
      return;

    case 'auto-provision':
      await autoProvision(deps, outcome.aadObjectId, outcome.displayName, reply);
      await deps.stateStore.save(deps.state);
      return;

    case 'resolved':
    case 'fallback': {
      const user = outcome.user;
      const key = `${deps.bot.appId}:${user.participantId}` +
        (outcome.kind === 'fallback' ? `:${user.aadObjectId}` : '');
      upsertConversation(
        deps.state,
        key,
        user.aadObjectId,
        user.displayName,
        activity,
      );
      await deps.stateStore.save(deps.state);
      deps.ctx.registry.setActive(user.participantId, 'teams');

      const text = activity.text?.trim();
      if (!text) {
        await reply('Empty message.');
        return;
      }

      if (text.startsWith('/')) {
        const handled = await handleCommand({
          appId: deps.bot.appId,
          participantId: user.participantId,
          aadObjectId: user.aadObjectId,
          conversationKey: key,
          isFallback: outcome.kind === 'fallback',
          text,
          state: deps.state,
          ctx: deps.ctx,
          options: deps.options,
          send: reply,
        });
        if (handled) {
          await deps.stateStore.save(deps.state);
          return;
        }
      }

      let typing: NodeJS.Timeout | undefined;
      const stopTyping = () => {
        if (typing) {
          clearInterval(typing);
          typing = undefined;
        }
      };
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
          conversationId: deps.state.legionConversations[key],
          options: deps.options,
          stream: context.stream,
          send: (activity) => context.send(activity),
          stopTyping,
        });
        if (conversationId !== undefined) {
          deps.state.legionConversations[key] = conversationId;
          await deps.stateStore.save(deps.state);
        }
      } finally {
        stopTyping();
      }
      return;
    }
  }
}
