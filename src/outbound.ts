// src/outbound.ts
import type { CollectiveLike } from './directory.js';
import { MessageActivityInput } from '@microsoft/teams.api';
import type { ResolvedOptions } from './options.js';
import { chunkMarkdown } from './render.js';
import { describeError } from './errors.js';
import type { StateStore, TeamsState } from './state.js';

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
  message: { senderId: string; recipientId: string; content: string },
): Promise<void> {
  const sender = deps.collective.get(message.senderId);
  const botIdentity = sender?.identities?.find((i) => i.connector === 'teams-bot');
  if (!botIdentity || deps.apps.ownerOf(botIdentity.externalId) !== message.senderId) return;

  const app = deps.apps.getByAppId(botIdentity.externalId);
  if (!app) return;

  const key = `${botIdentity.externalId}:${message.recipientId}`;
  let record = deps.state.conversations[key];

  if (!record) {
    const recipient = deps.collective.get(message.recipientId);
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
        tenantId: deps.options.tenantId,
        aadObjectId: userIdentity.externalId,
        updatedAt: new Date().toISOString(),
      };
      deps.state.conversations[key] = record;
      await deps.stateStore.save(deps.state);
    } else {
      for (const [candidateKey, candidate] of Object.entries(deps.state.conversations)) {
        if (candidateKey.startsWith(`${key}:`) && (!record || candidate.updatedAt > record.updatedAt)) {
          record = candidate;
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
