// src/inbound.ts
import type { ConnectorContext } from '@legion-collective/core';
import type { CollectiveLike } from './directory.js';
import { MessageActivityInput, TypingActivityInput, stripMentionsText, type IMessageActivity } from '@microsoft/teams.api';
import { applyGroupPrefix } from './attribution.js';
import { handleCommand } from './commands.js';
import { isAwaitingApproval } from './approvals.js';
import { runCommunicate, type TeamsStreamer } from './communicate.js';
import type { BotEntry } from './directory.js';
import {
  findTeamsActor,
  hash6,
  provisionedGroupId,
  provisionedUserId,
  readTeamsId,
  resolveGroupIdentity,
  resolveIdentity,
  type ResolvedGroup,
} from './identity.js';
import type { ResolvedOptions } from './options.js';
import { sendText } from './render.js';
import type { StateStore, TeamsState } from './state.js';
import { recordMessage } from './tool-state.js';

export interface TeamsActivity {
  id?: string;
  from?: { id?: string; name?: string; aadObjectId?: string };
  recipient?: { id?: string };
  conversation?: { id?: string; name?: string; conversationType?: string };
  channelData?: { tenant?: { id?: string } };
  serviceUrl?: string;
  replyToId?: string;
  text?: string;
  entities?: Array<{
    type?: string;
    text?: string;
    mentioned?: { id?: string; name?: string };
  }>;
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

function upsertPersonalConversation(
  state: TeamsState,
  key: string,
  externalId: string,
  displayName: string | undefined,
  activity: TeamsActivity,
): void {
  state.conversations[key] = {
    conversationId: activity.conversation?.id ?? '',
    conversationType: 'personal',
    externalId,
    tenantId: activity.channelData?.tenant?.id,
    serviceUrl: activity.serviceUrl,
    displayName,
    updatedAt: new Date().toISOString(),
  };
}

function upsertGroupConversation(
  state: TeamsState,
  key: string,
  group: ResolvedGroup,
  activity: TeamsActivity,
): void {
  state.conversations[key] = {
    conversationId: group.conversationId,
    conversationType: 'groupChat',
    externalId: group.conversationId,
    name: group.name,
    tenantId: activity.channelData?.tenant?.id,
    serviceUrl: activity.serviceUrl,
    updatedAt: new Date().toISOString(),
  };
}

function stripBotMention(activity: TeamsActivity): string {
  if (!activity.recipient?.id) return activity.text?.trim() ?? '';
  return stripMentionsText(activity as unknown as IMessageActivity, {
    accountId: activity.recipient.id,
  })?.trim() ?? '';
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

async function autoProvisionGroup(
  deps: InboundDeps,
  conversationId: string,
  conversationName: string | undefined,
  reply: (text: string) => Promise<void>,
): Promise<void> {
  const id = provisionedGroupId(conversationName, conversationId);
  const name = conversationName?.trim() || `Teams group ${hash6(conversationId)}`;
  const existing = deps.collective.get(id);

  if (existing && (existing.type !== 'user' || existing.status === 'retired' || existing.identities?.some(
    (identity) => identity.connector === 'teams-group' && identity.externalId !== conversationId,
  ))) {
    await reply(`Participant ${id} is already in use; ask an administrator to pair this group.`);
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
      { participantId: id, connector: 'teams-group', externalId: conversationId },
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

  await reply(`Welcome, ${name}. This group is now paired as ${id}.`);
}

async function runTurn(args: {
  deps: InboundDeps;
  context: TeamsContextLike;
  participantId: string;
  key: string;
  text: string;
}): Promise<void> {
  const { deps, context, participantId, key, text } = args;
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
    const turn = await runCommunicate({
      ctx: deps.ctx,
      userId: participantId,
      botOwnerId: deps.bot.participantId,
      text,
      conversationId: deps.state.legionConversations[key],
      options: deps.options,
      stream: context.stream,
      send: (activity) => context.send(activity),
      stopTyping,
    });

    const { conversationId, storedMessageId, outboundActivityId } = turn;
    let changed = false;
    if (conversationId !== undefined) {
      deps.state.legionConversations[key] = conversationId;
      deps.state.legionConversationTargets[conversationId] = key;
      changed = true;
    }

    const teamsConversationId = context.activity.conversation?.id;
    const inboundActivityId = context.activity.id;
    if (teamsConversationId && inboundActivityId && storedMessageId) {
      recordMessage(
        deps.state,
        {
          appId: deps.bot.appId,
          conversationId: teamsConversationId,
          activityId: inboundActivityId,
          direction: 'inbound',
          legionMessageId: storedMessageId,
          ...(conversationId === undefined ? {} : { legionConversationId: conversationId }),
        },
        deps.options.maxMessageRecords,
      );
      changed = true;
    }
    if (teamsConversationId && outboundActivityId) {
      recordMessage(
        deps.state,
        {
          appId: deps.bot.appId,
          conversationId: teamsConversationId,
          activityId: outboundActivityId,
          direction: 'outbound',
          ...(conversationId === undefined ? {} : { legionConversationId: conversationId }),
        },
        deps.options.maxMessageRecords,
      );
      changed = true;
    }
    if (changed) await deps.stateStore.save(deps.state);
  } finally {
    stopTyping();
  }
}

export async function handleInbound(
  deps: InboundDeps,
  context: TeamsContextLike,
): Promise<void> {
  const activity = context.activity;
  const reply = (text: string) => sendText((message) => context.send(message), deps.options.aiLabel, text);
  const conversationType = activity.conversation?.conversationType;
  const isGroup = conversationType === 'groupChat';

  if (conversationType !== 'personal' && !(isGroup && deps.options.groupChats)) {
    await reply(deps.options.groupChats
      ? 'Only personal chats and group chats are supported.'
      : 'Only personal chats are supported.');
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

  if (isGroup) {
    const conversationId = activity.conversation?.id?.trim();
    if (!conversationId) {
      await reply('Unable to read the group conversation.');
      return;
    }

    const outcome = resolveGroupIdentity({
      collective: deps.collective,
      state: deps.state,
      appId: deps.bot.appId,
      conversationId,
      name: activity.conversation?.name,
      options: deps.options,
    });

    switch (outcome.kind) {
      case 'pairing':
        await deps.stateStore.save(deps.state);
        await reply(
          `This group isn't registered with Legion yet. Ask someone with the \`pair\` tool to run ` +
          `/pair ${outcome.code} <participant-id>. Group ID: ${conversationId}.`,
        );
        return;

      case 'reject':
        await reply(
          `This group isn't registered with Legion yet. Group ID: ${conversationId}. ` +
          'Ask an administrator to register it.',
        );
        return;

      case 'auto-provision':
        await autoProvisionGroup(deps, conversationId, outcome.name, reply);
        await deps.stateStore.save(deps.state);
        return;

      case 'resolved':
      case 'fallback': {
        const group = outcome.group;
        const key = `${deps.bot.appId}:${group.participantId}:${conversationId}`;
        upsertGroupConversation(deps.state, key, group, activity);
        await deps.stateStore.save(deps.state);
        deps.ctx.registry.setActive(group.participantId, 'teams');

        const text = stripBotMention(activity);
        if (text.startsWith('/')) {
          const sender = findTeamsActor(deps.collective, aadObjectId);
          const handled = await handleCommand({
            appId: deps.bot.appId,
            participantId: group.participantId,
            senderParticipantId: sender?.id,
            aadObjectId,
            senderName: activity.from?.name,
            conversationKey: key,
            conversationId,
            isFallback: outcome.kind === 'fallback',
            isGroup: true,
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

        const message = text
          ? applyGroupPrefix(deps.options.groupSenderPrefix, activity.from?.name ?? 'Unknown', text)
          : '';
        await runTurn({
          deps,
          context,
          participantId: group.participantId,
          key,
          text: message,
        });
        return;
      }
    }
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
      upsertPersonalConversation(deps.state, key, user.aadObjectId, user.displayName, activity);
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
          senderParticipantId: user.participantId,
          aadObjectId: user.aadObjectId,
          senderName: user.displayName,
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

      await runTurn({
        deps,
        context,
        participantId: user.participantId,
        key,
        text,
      });
      return;
    }
  }
}
