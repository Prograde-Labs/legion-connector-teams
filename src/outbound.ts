// src/outbound.ts
import type { CollectiveLike } from './directory.js';
import { MessageActivityInput } from '@microsoft/teams.api';
import type { ResolvedOptions } from './options.js';
import { chunkMarkdown } from './render.js';
import { describeError } from './errors.js';
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
  message: { senderId: string; recipientId: string; conversationId: string; content: string },
): Promise<void> {
  const sender = deps.collective.get(message.senderId);
  const botIdentity = sender?.identities?.find((i) => i.connector === 'teams-bot');
  if (!botIdentity || deps.apps.ownerOf(botIdentity.externalId) !== message.senderId) return;

  const app = deps.apps.getByAppId(botIdentity.externalId);
  if (!app) return;

  const appId = botIdentity.externalId;
  const recipient = deps.collective.get(message.recipientId);
  let record: ConversationRecord | undefined;

  const targetKey = deps.state.legionConversationTargets[message.conversationId];
  if (targetKey) record = deps.state.conversations[targetKey];

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
      deps.state.legionConversationTargets[message.conversationId] = key;
      await deps.stateStore.save(deps.state);
    } else {
      const key = `${appId}:${message.recipientId}`;
      record = deps.state.conversations[key];

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
            console.error('Teams conversation creation failed', describeError(error));
            return;
          }
          record = {
            conversationId: conversation.id,
            conversationType: 'personal',
            externalId: userIdentity.externalId,
            tenantId: deps.options.tenantId,
            updatedAt: new Date().toISOString(),
          };
          deps.state.conversations[key] = record;
          deps.state.legionConversationTargets[message.conversationId] = key;
          await deps.stateStore.save(deps.state);
        } else {
          for (const [candidateKey, candidate] of Object.entries(deps.state.conversations)) {
            if (candidateKey.startsWith(`${key}:`) && (!record || candidate.updatedAt > record.updatedAt)) {
              record = candidate;
            }
          }
        }
      }
    }
  }
  if (!record) return;

  for (const chunk of chunkMarkdown(message.content, deps.options.maxMessageChars)) {
    const activity = new MessageActivityInput(chunk);
    if (deps.options.aiLabel) activity.addAiGenerated();
    try {
      await app.send(record.conversationId, activity);
    } catch (error) {
      console.error('Teams message delivery failed', describeError(error));
      return;
    }
  }
}
