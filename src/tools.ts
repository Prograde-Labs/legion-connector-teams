import type { AnyTool, Tool, ToolContext, ToolResult } from '@legion-collective/core';
import type { IAdaptiveCard } from '@microsoft/teams.cards';
import { MessageActivityInput } from '@microsoft/teams.api';
import { describeError } from './errors.js';
import {
  resolveCallerBot,
  resolveConversationTarget,
  resolveMessageTarget,
  type TeamsToolDeps,
} from './tool-targets.js';
import { toolButtonActions, type ToolButtonSpec } from './render.js';
import { recordMessage } from './tool-state.js';

const ok = (data: unknown): ToolResult => ({ status: 'success', data });
const fail = (error: string): ToolResult => ({ status: 'error', error });

const MAX_CARD_ACTIONS = 6;
const MAX_BUTTON_TITLE = 40;

function validateButtons(
  value: unknown,
): { ok: true; buttons: ToolButtonSpec[] } | { ok: false; error: string } {
  if (value === undefined) return { ok: true, buttons: [] };
  if (!Array.isArray(value)) return { ok: false, error: 'buttons must be an array' };
  if (value.length === 0) return { ok: true, buttons: [] };
  if (value.length > MAX_CARD_ACTIONS) {
    return { ok: false, error: `buttons must contain at most ${MAX_CARD_ACTIONS} entries` };
  }
  const buttons: ToolButtonSpec[] = [];
  for (const [index, raw] of value.entries()) {
    if (typeof raw !== 'object' || raw === null || Array.isArray(raw)) {
      return { ok: false, error: `buttons[${index}] must be an object` };
    }
    const entry = raw as Record<string, unknown>;
    const title = typeof entry.title === 'string' ? entry.title.trim() : '';
    if (!title) return { ok: false, error: `buttons[${index}].title is required` };
    if (title.length > MAX_BUTTON_TITLE) {
      return {
        ok: false,
        error: `buttons[${index}].title must be at most ${MAX_BUTTON_TITLE} characters`,
      };
    }
    const tool = typeof entry.tool === 'string' ? entry.tool.trim() : '';
    if (!tool) return { ok: false, error: `buttons[${index}].tool is required` };
    if (
      entry.args !== undefined &&
      (typeof entry.args !== 'object' || entry.args === null || Array.isArray(entry.args))
    ) {
      return { ok: false, error: `buttons[${index}].args must be an object` };
    }
    if (
      entry.resultCard !== undefined &&
      (typeof entry.resultCard !== 'object' ||
        entry.resultCard === null ||
        Array.isArray(entry.resultCard))
    ) {
      return { ok: false, error: `buttons[${index}].resultCard must be an object` };
    }
    buttons.push({
      title,
      tool,
      ...(entry.args === undefined ? {} : { args: entry.args }),
      ...(entry.resultCard === undefined ? {} : { resultCard: entry.resultCard }),
    });
  }
  return { ok: true, buttons };
}

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
    description:
      "Send an Adaptive Card to a Teams conversation through the caller's Teams bot. " +
      'Optional buttons run granted tools as the clicking user.',
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
        buttons: {
          type: 'array',
          description: 'Optional tool buttons. Each runs a granted tool as the clicking user.',
          items: {
            type: 'object',
            properties: {
              title: { type: 'string' },
              tool: { type: 'string' },
              args: { type: 'object' },
              resultCard: { type: 'object' },
            },
            required: ['title', 'tool'],
          },
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

      const checked = validateButtons(raw.buttons);
      if (!checked.ok) return fail(`teams__send_card: ${checked.error}`);

      const existingActions = Array.isArray((card as { actions?: unknown[] }).actions)
        ? (card as { actions: unknown[] }).actions
        : [];
      const actions = [...existingActions, ...toolButtonActions(checked.buttons)];
      if (checked.buttons.length > 0 && actions.length > MAX_CARD_ACTIONS) {
        return fail(
          `teams__send_card: card has ${actions.length} actions (at most ${MAX_CARD_ACTIONS} actions)`,
        );
      }
      const outgoingCard = actions.length
        ? { ...(card as Record<string, unknown>), actions }
        : card;

      const serialized = JSON.stringify(outgoingCard);
      const bytes = Buffer.byteLength(serialized, 'utf8');
      if (bytes > deps.options.maxCardBytes) {
        return fail(`teams__send_card: card is ${bytes} bytes (max ${deps.options.maxCardBytes})`);
      }

      const activity =
        typeof raw.fallbackText === 'string' && raw.fallbackText.trim()
          ? new MessageActivityInput('', { summary: raw.fallbackText })
          : new MessageActivityInput();
      activity.addCard('adaptive', outgoingCard as IAdaptiveCard);

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

export function postReactionTool(deps: TeamsToolDeps): Tool {
  return {
    name: 'teams__post_reaction',
    description:
      "Add or remove an emoji reaction on a Teams message through the caller's Teams bot.",
    parameters: {
      type: 'object',
      properties: {
        messageId: {
          type: 'string',
          description: 'Legion message id or Teams activity id',
        },
        reaction: {
          type: 'string',
          description: 'Emoji reaction type, e.g. like, heart, 1f440_eyes, launch',
        },
        conversationId: {
          type: 'string',
          description:
            'Teams conversation id; required when messageId is a raw activity id that is not recorded',
        },
        remove: {
          type: 'boolean',
          description: 'Remove instead of add',
        },
      },
      required: ['messageId', 'reaction'],
    },
    async execute(args: unknown, context: ToolContext): Promise<ToolResult> {
      const raw = (args ?? {}) as Record<string, unknown>;
      const messageId = typeof raw.messageId === 'string' ? raw.messageId.trim() : '';
      const reaction = typeof raw.reaction === 'string' ? raw.reaction.trim() : '';
      if (!messageId) return fail('teams__post_reaction: messageId is required');
      if (!reaction) return fail('teams__post_reaction: reaction is required');

      const conversationId =
        typeof raw.conversationId === 'string' && raw.conversationId.trim()
          ? raw.conversationId.trim()
          : undefined;

      const caller = resolveCallerBot(deps, context.participant.id);
      if (!caller.ok) return fail(caller.error);

      const target = resolveMessageTarget(deps, caller.value.appId, {
        messageId,
        conversationId,
      });
      if (!target.ok) return fail(target.error);

      const remove = raw.remove === true;
      try {
        if (remove) {
          await caller.value.app.api.reactions.delete(
            target.value.conversationId,
            target.value.activityId,
            reaction,
          );
        } else {
          await caller.value.app.api.reactions.add(
            target.value.conversationId,
            target.value.activityId,
            reaction,
          );
        }
      } catch (error) {
        return fail(`teams__post_reaction: ${describeError(error)}`);
      }

      return ok({
        conversationId: target.value.conversationId,
        activityId: target.value.activityId,
        reaction,
        removed: remove,
      });
    },
  };
}

export function listConversationsTool(deps: TeamsToolDeps): Tool {
  return {
    name: 'teams__list_conversations',
    description: "List Teams conversations known to the caller's Teams bot.",
    parameters: {
      type: 'object',
      properties: {
        conversationType: {
          type: 'string',
          enum: ['personal', 'groupChat'],
        },
        limit: {
          type: 'integer',
          minimum: 1,
          maximum: 200,
        },
      },
    },
    async execute(args: unknown, context: ToolContext): Promise<ToolResult> {
      const raw = (args ?? {}) as Record<string, unknown>;
      const type = raw.conversationType;
      if (type !== undefined && type !== 'personal' && type !== 'groupChat') {
        return fail(`teams__list_conversations: invalid conversationType '${String(type)}'`);
      }
      const limit =
        typeof raw.limit === 'number' && Number.isInteger(raw.limit)
          ? Math.min(200, Math.max(1, raw.limit))
          : 50;

      const caller = resolveCallerBot(deps, context.participant.id);
      if (!caller.ok) return fail(caller.error);

      const prefix = `${caller.value.appId}:`;
      const conversations = Object.entries(deps.state.conversations)
        .filter(
          ([key, record]) =>
            key.startsWith(prefix) && (type === undefined || record.conversationType === type),
        )
        .sort(([, left], [, right]) => right.updatedAt.localeCompare(left.updatedAt))
        .slice(0, limit)
        .map(([key, record]) => ({
          key,
          conversationId: record.conversationId,
          conversationType: record.conversationType,
          ...(record.externalId === undefined ? {} : { externalId: record.externalId }),
          ...(record.name === undefined ? {} : { name: record.name }),
          ...(record.displayName === undefined ? {} : { displayName: record.displayName }),
          updatedAt: record.updatedAt,
        }));

      return ok({ conversations });
    },
  };
}

export function buildTeamsTools(deps: TeamsToolDeps): AnyTool[] {
  const tools: AnyTool[] = [];
  if (deps.options.tools.has('cards')) tools.push(sendCardTool(deps));
  if (deps.options.tools.has('reactions')) tools.push(postReactionTool(deps));
  if (deps.options.tools.has('list')) tools.push(listConversationsTool(deps));
  return tools;
}
