import type { AnyTool, Tool, ToolContext, ToolResult } from '@legion-collective/core';
import type { IAdaptiveCard } from '@microsoft/teams.cards';
import { MessageActivityInput } from '@microsoft/teams.api';
import { describeError } from './errors.js';
import { resolveCallerBot, resolveConversationTarget, type TeamsToolDeps } from './tool-targets.js';
import { recordMessage } from './tool-state.js';

const ok = (data: unknown): ToolResult => ({ status: 'success', data });
const fail = (error: string): ToolResult => ({ status: 'error', error });

async function saveQuietly(deps: TeamsToolDeps): Promise<void> {
  try {
    await deps.store.save(deps.state);
  } catch (error) {
    console.error('Teams state save failed', describeError(error));
  }
}

export function sendCardTool(deps: TeamsToolDeps): Tool {
  return {
    name: 'teams__send_card',
    description: "Send an Adaptive Card to a Teams conversation through the caller's Teams bot.",
    parameters: {
      type: 'object',
      properties: {
        participantId: {
          type: 'string',
          description: 'Legion user or group participant to send to',
        },
        conversationId: {
          type: 'string',
          description: "Raw Teams conversation id; must belong to the caller's bot",
        },
        card: {
          type: 'object',
          description: 'Adaptive Card JSON (schema version 1.4 or later)',
        },
        fallbackText: {
          type: 'string',
          description: 'Plain-text summary shown when the card cannot render',
        },
        aiLabel: {
          type: 'boolean',
          description: 'Mark the message as AI-generated',
        },
      },
      required: ['card'],
    },
    async execute(args: unknown, context: ToolContext): Promise<ToolResult> {
      const raw = (args ?? {}) as Record<string, unknown>;
      const card = raw.card;
      if (typeof card !== 'object' || card === null || Array.isArray(card)) {
        return fail('teams__send_card: card must be an object');
      }

      const participantId =
        typeof raw.participantId === 'string' && raw.participantId.trim()
          ? raw.participantId.trim()
          : undefined;
      const conversationId =
        typeof raw.conversationId === 'string' && raw.conversationId.trim()
          ? raw.conversationId.trim()
          : undefined;
      if ((participantId === undefined) === (conversationId === undefined)) {
        return fail('teams__send_card: exactly one of participantId or conversationId is required');
      }

      const caller = resolveCallerBot(deps, context.participant.id);
      if (!caller.ok) return fail(caller.error);

      const target = await resolveConversationTarget(deps, caller.value.appId, caller.value.app, {
        participantId,
        conversationId,
      });
      if (!target.ok) return fail(target.error);

      const serialized = JSON.stringify(card);
      const bytes = Buffer.byteLength(serialized, 'utf8');
      if (bytes > deps.options.maxCardBytes) {
        return fail(`teams__send_card: card is ${bytes} bytes (max ${deps.options.maxCardBytes})`);
      }

      const activity =
        typeof raw.fallbackText === 'string' && raw.fallbackText.trim()
          ? new MessageActivityInput('', { summary: raw.fallbackText })
          : new MessageActivityInput();
      activity.addCard('adaptive', card as IAdaptiveCard);

      const aiLabel = typeof raw.aiLabel === 'boolean' ? raw.aiLabel : deps.options.aiLabel;
      if (aiLabel) activity.addAiGenerated();

      let sent: { id?: string };
      try {
        sent = await caller.value.app.send(target.value.conversationId, activity);
      } catch (error) {
        return fail(`teams__send_card: ${describeError(error)}`);
      }

      const activityId = sent?.id;
      if (activityId) {
        recordMessage(
          deps.state,
          {
            appId: caller.value.appId,
            conversationId: target.value.conversationId,
            activityId,
            direction: 'outbound',
          },
          deps.options.maxMessageRecords,
        );
      }
      await saveQuietly(deps);

      return ok({
        conversationId: target.value.conversationId,
        ...(activityId === undefined ? {} : { activityId }),
      });
    },
  };
}

export function buildTeamsTools(deps: TeamsToolDeps): AnyTool[] {
  const tools: AnyTool[] = [];
  if (deps.options.tools.has('cards')) tools.push(sendCardTool(deps));
  return tools;
}
