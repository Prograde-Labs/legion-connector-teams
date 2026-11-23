// src/outbound.ts
import type { ApprovalRequestView, DeliveryOutcome } from '@legion-collective/core';
import type { CollectiveLike } from './directory.js';
import { MessageActivityInput } from '@microsoft/teams.api';
import type { ResolvedOptions } from './options.js';
import { chunkMarkdown } from './render.js';
import { describeError } from './errors.js';
import { sendApprovalCards } from './approvals.js';
import type { ConversationRecord, StateStore, TeamsState } from './state.js';

export interface OutboundApp {
  send(conversationId: string, activity: MessageActivityInput): Promise<unknown>;
  api: {
    conversations: {
      create(params: Record<string, unknown>): Promise<{ id: string }>;
    };
  };
}

export interface OutboundDeps {
  collective: CollectiveLike;
  state: TeamsState;
  stateStore: StateStore;
  apps: { getByAppId(appId: string): OutboundApp | undefined; ownerOf(appId: string): string | undefined };
  options: ResolvedOptions;
}

export async function deliverMessage(
  deps: OutboundDeps,
  message: {
    senderId: string;
    recipientId: string;
    conversationId: string;
    replyTo?: string;
    content: string;
    approvalRequests?: ApprovalRequestView[];
  },
): Promise<DeliveryOutcome> {
  const sender = deps.collective.get(message.senderId);
  const botOwner = sender?.identities?.some((i) => i.connector === 'teams-bot')
    ? sender
    : message.replyTo
      ? deps.collective.get(message.replyTo)
      : undefined;
  const botIdentity = botOwner?.identities?.find((i) => i.connector === 'teams-bot');
  if (!botIdentity || !botOwner || deps.apps.ownerOf(botIdentity.externalId) !== botOwner.id) {
    return {
      status: 'not-reachable',
      reason: `sender '${message.senderId}' has no live Teams bot`,
    };
  }

  const app = deps.apps.getByAppId(botIdentity.externalId);
  if (!app) {
    return {
      status: 'not-reachable',
      reason: `Teams app '${botIdentity.externalId}' is not running`,
    };
  }

  const appId = botIdentity.externalId;
  const recipient = deps.collective.get(message.recipientId);
  let record: ConversationRecord | undefined;
  let recordKey: string | undefined;

  const mappedKey = deps.state.legionConversationTargets[message.conversationId];
  if (
    mappedKey &&
    (mappedKey === `${appId}:${message.recipientId}` ||
      mappedKey.startsWith(`${appId}:${message.recipientId}:`))
  ) {
    record = deps.state.conversations[mappedKey];
    if (record) recordKey = mappedKey;
  }

  if (!record) {
    const groupIdentity = recipient?.identities?.find((i) => i.connector === 'teams-group');
    if (groupIdentity) {
      const key = `${appId}:${message.recipientId}:${groupIdentity.externalId}`;
      record = deps.state.conversations[key] ?? {
        conversationId: groupIdentity.externalId,
        conversationType: 'groupChat',
        externalId: groupIdentity.externalId,
        updatedAt: new Date().toISOString(),
      };
      deps.state.conversations[key] = record;
      recordKey = key;
    } else {
      const key = `${appId}:${message.recipientId}`;
      record = deps.state.conversations[key];
      recordKey = key;

      if (!record) {
        const userIdentity = recipient?.identities?.find((i) => i.connector === 'teams');
        if (userIdentity) {
          let conversation: { id: string };
          try {
            conversation = await app.api.conversations.create({
              members: [{ id: userIdentity.externalId }],
              tenantId: deps.options.tenantId,
              channelData: { tenant: { id: deps.options.tenantId } },
            });
          } catch (error) {
            const reason = describeError(error);
            console.error('Teams conversation creation failed', reason);
            return {
              status: 'failed',
              reason: `Teams conversation creation failed: ${reason}`,
            };
          }
          record = {
            conversationId: conversation.id,
            conversationType: 'personal',
            externalId: userIdentity.externalId,
            tenantId: deps.options.tenantId,
            updatedAt: new Date().toISOString(),
          };
          deps.state.conversations[key] = record;
        } else {
          for (const [candidateKey, candidate] of Object.entries(deps.state.conversations)) {
            if (
              candidateKey.startsWith(`${key}:`) &&
              (candidate.conversationType ?? 'personal') === 'personal' &&
              (!record || candidate.updatedAt > record.updatedAt)
            ) {
              record = candidate;
              recordKey = candidateKey;
            }
          }
        }
      }
    }
  }

  if (!record || !recordKey) {
    return {
      status: 'not-reachable',
      reason: `no Teams conversation for '${message.recipientId}'`,
    };
  }

  for (const chunk of chunkMarkdown(message.content, deps.options.maxMessageChars)) {
    const activity = new MessageActivityInput(chunk);
    if (deps.options.aiLabel) activity.addAiGenerated();
    try {
      await app.send(record.conversationId, activity);
    } catch (error) {
      const reason = describeError(error);
      console.error('Teams message delivery failed', reason);
      return { status: 'failed', reason: `Teams message delivery failed: ${reason}` };
    }
  }

  if (message.approvalRequests?.length) {
    try {
      await sendApprovalCards(message.approvalRequests, (activity) =>
        app.send(record!.conversationId, activity),
      );
    } catch (error) {
      console.error('Teams approval card delivery failed', describeError(error));
    }
  }

  deps.state.legionConversationTargets[message.conversationId] = recordKey;
  if (deps.options.learnOutboundThread) {
    deps.state.legionConversations[recordKey] = message.conversationId;
  }
  try {
    await deps.stateStore.save(deps.state);
  } catch (error) {
    console.error('Teams state save failed after delivery', describeError(error));
  }

  return { status: 'delivered' };
}
