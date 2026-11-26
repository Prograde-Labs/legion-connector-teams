// src/commands.ts
import type { ConnectorContext } from '@legion-collective/core';
import { isAwaitingApproval } from './approvals.js';
import { validatePendingPair } from './pairing.js';
import {
  bindConversation,
  conversationCard,
  conversationErrorText,
  listConversations,
} from './conversations.js';
import type { ConversationEntry } from './conversations.js';
import type { ResolvedOptions } from './options.js';
import { formatToolResult } from './render.js';
import type { TeamsState } from './state.js';
import type { IAdaptiveCard } from '@microsoft/teams.cards';

const pairingInFlight = new Set<string>();

export interface CommandDeps {
  appId: string;
  participantId: string;
  senderParticipantId?: string;
  aadObjectId: string;
  senderName?: string;
  conversationKey?: string;
  conversationId?: string;
  isFallback?: boolean;
  isGroup?: boolean;
  text: string;
  state: TeamsState;
  ctx: ConnectorContext;
  options: ResolvedOptions;
  send: (text: string) => Promise<void>;
  sendCard: (card: IAdaptiveCard) => Promise<void>;
  stopTyping?: () => void;
}

function toolFailure(result: { status: string; error?: string; message?: string }): string {
  if (result.status === 'error') return result.error ?? 'unknown error';
  return formatToolResult(result) || result.status;
}

function formatCount(value: unknown): string {
  const number = typeof value === 'number' && Number.isFinite(value) ? value : 0;
  return number.toLocaleString('en-US');
}

function formatUsage(totals: {
  input?: unknown;
  output?: unknown;
  reasoning?: unknown;
  cache?: { read?: unknown; write?: unknown };
  cost?: unknown;
}): string {
  const cost = typeof totals.cost === 'number' && Number.isFinite(totals.cost) ? totals.cost : 0;
  return [
    `Tokens: in ${formatCount(totals.input)} · out ${formatCount(totals.output)}` +
      ` · reasoning ${formatCount(totals.reasoning)}`,
    `Cache: ${formatCount(totals.cache?.read)} read / ${formatCount(totals.cache?.write)} write` +
      ` · Cost: $${cost.toFixed(4)}`,
  ].join('\n');
}

export async function handleCommand(deps: CommandDeps): Promise<boolean> {
  const [raw, ...rest] = deps.text.trim().split(/\s+/);
  const command = raw.toLowerCase();
  const key = deps.conversationKey ?? `${deps.appId}:${deps.participantId}`;

  switch (command) {
    case '/new': {
      const previous = deps.state.legionConversations[key];
      delete deps.state.legionConversations[key];
      if (previous) delete deps.state.legionConversationTargets[previous];
      await deps.send('Started a new conversation.');
      return true;
    }
    case '/pair': {
      if (deps.isFallback && !(deps.isGroup && deps.senderParticipantId)) {
        await deps.send('Only registered users can redeem pairing codes.');
        return true;
      }
      const callerId = deps.senderParticipantId ?? (deps.isGroup ? undefined : deps.participantId);
      if (!callerId) {
        await deps.send('Only registered users can redeem pairing codes.');
        return true;
      }
      const code = rest[0]?.toUpperCase();
      const target = rest[1];
      if (!code || !target) {
        await deps.send('Usage: /pair <code> <participant-id>');
        return true;
      }
      if (pairingInFlight.has(code)) {
        await deps.send('Pairing already in progress.');
        return true;
      }
      pairingInFlight.add(code);
      try {
        const pending = validatePendingPair(deps.state, code);
        if (!pending || pending.appId !== deps.appId) {
          await deps.send('Pairing code is invalid or expired.');
          return true;
        }
        const caller = deps.options.pairingParticipantId ?? callerId;
        const { result } = await deps.ctx.callTool(caller, 'pair', {
          participantId: target,
          connector: pending.connector,
          externalId: pending.externalId,
        });
        if (result.status === 'success') {
          delete deps.state.pendingPairs[code.toUpperCase()];
          const label = pending.connector === 'teams'
            ? 'Teams'
            : pending.connector === 'teams-group'
              ? 'Teams group'
              : pending.connector;
          await deps.send(`Paired ${target} with ${label} identity ${pending.externalId}.`);
        } else if (isAwaitingApproval(result)) {
          await deps.send('Pairing is awaiting approval.');
        } else {
          await deps.send(`Pairing failed: ${result.error ?? result.status}`);
        }
        return true;
      } finally {
        pairingInFlight.delete(code);
      }
    }
    case '/whoami': {
      const conversation = deps.state.legionConversations[key] ?? '(none)';
      if (deps.isGroup) {
        await deps.send(
          `Participant: ${deps.participantId}\n` +
          `Group conversation ID: ${deps.conversationId ?? '(unknown)'}\n` +
          `Sender: ${deps.senderName ?? deps.aadObjectId}\n` +
          `Bot app: ${deps.appId}\nConversation: ${conversation}`,
        );
      } else {
        await deps.send(
          `Participant: ${deps.participantId}\nTeams ID: ${deps.aadObjectId}\n` +
          `Bot app: ${deps.appId}\nConversation: ${conversation}`,
        );
      }
      return true;
    }
    case '/help': {
      await deps.send(
        'Commands:\n' +
        '/new — start a new conversation\n' +
        '/pair <code> <participant-id> — pair a pending Teams identity\n' +
        '/whoami — show your identity\n' +
        'In a group chat, /new and /whoami act on the group; /pair runs as you.\n' +
        '/help — this message',
      );
      return true;
    }
    case '/switch': {
      const currentId = deps.state.legionConversations[key];
      try {
        const entries = await listConversations(deps.ctx, deps.participantId);
        const target = rest[0];
        if (!target) {
          await deps.sendCard(conversationCard(entries, currentId, 0, key));
          return true;
        }
        const entry = entries.find((candidate) => candidate.id === target);
        if (!entry) {
          await deps.send('Conversation not found. Use /switch to list active conversations.');
          return true;
        }
        if (entry.id === currentId) {
          await deps.send(`Already on ${entry.title}.`);
          return true;
        }
        bindConversation(deps.state, key, entry.id);
        await deps.send(`Switched to ${entry.title} (${entry.messageCount} messages).`);
        return true;
      } catch (error) {
        await deps.send(`Unable to list conversations: ${conversationErrorText(error)}`);
        return true;
      }
    }
    case '/compact': {
      const compaction = deps.options.compaction;
      if (!compaction) {
        await deps.send('Compaction is not configured. Set the "compaction.agentId" connector option.');
        return true;
      }
      const conversationId = deps.state.legionConversations[key];
      if (!conversationId) {
        await deps.send('Nothing to compact — no conversation yet.');
        return true;
      }
      try {
        const loaded = await deps.ctx.callTool(deps.participantId, 'get_conversation', {
          conversationId,
        });
        if (loaded.result.status !== 'success') {
          await deps.send(`Compaction failed: ${toolFailure(loaded.result)}`);
          return true;
        }
        const messages = (loaded.result.data as { messages?: Array<{ id?: unknown }> } | undefined)
          ?.messages ?? [];
        const messageIds = messages
          .map((message) => message.id)
          .filter((id): id is string => typeof id === 'string' && id.length > 0);
        if (messageIds.length === 0) {
          await deps.send('Nothing to compact — conversation is empty.');
          return true;
        }
        const instruction = rest.join(' ').trim() || compaction.instruction;
        const compacted = await deps.ctx.callTool(deps.participantId, 'compact_conversation', {
          conversationId,
          messageIds,
          agentId: compaction.agentId,
          ...(instruction ? { instruction } : {}),
        });
        if (compacted.result.status !== 'success') {
          await deps.send(`Compaction failed: ${toolFailure(compacted.result)}`);
          return true;
        }
        const summaryMessageId = (compacted.result.data as { summaryMessageId?: unknown } | undefined)
          ?.summaryMessageId;
        if (typeof summaryMessageId !== 'string' || !summaryMessageId) {
          await deps.send(`Compacted ${messageIds.length} messages — no summary message was returned.`);
          return true;
        }
        const refreshed = await deps.ctx.callTool(deps.participantId, 'get_conversation', {
          conversationId,
        });
        const summary = refreshed.result.status === 'success'
          ? ((refreshed.result.data as {
            messages?: Array<{ id?: unknown; content?: unknown }>;
          } | undefined)?.messages ?? []).find((message) => message.id === summaryMessageId)
          : undefined;
        const content = typeof summary?.content === 'string' ? summary.content.trim() : '';
        if (!content) {
          await deps.send(`Compacted ${messageIds.length} messages — summary unavailable.`);
          return true;
        }
        deps.stopTyping?.();
        await deps.send(
          `🗜 Compacted ${messageIds.length} messages into a summary:\n\n${content.slice(0, 3500)}`,
        );
        return true;
      } catch (error) {
        await deps.send(`Compaction failed: ${error instanceof Error ? error.message : String(error)}`);
        return true;
      }
    }
    case '/context': {
      const conversationId = deps.state.legionConversations[key];
      if (!conversationId) {
        await deps.send('No conversation yet — send a message first.');
        return true;
      }
      let loaded;
      try {
        loaded = await deps.ctx.callTool(deps.participantId, 'get_conversation', {
          conversationId,
        });
      } catch (error) {
        await deps.send(`Context unavailable: ${error instanceof Error ? error.message : String(error)}`);
        return true;
      }
      if (loaded.result.status !== 'success') {
        await deps.send(`Context unavailable: ${toolFailure(loaded.result)}`);
        return true;
      }
      const data = loaded.result.data as {
        title?: unknown;
        status?: unknown;
        messages?: unknown;
      } | undefined;
      const activeMessages = Array.isArray(data?.messages) ? data.messages.length : 0;
      const fallbackTitle =
        typeof data?.title === 'string' && data.title.trim() ? data.title.trim() : conversationId;
      const status =
        typeof data?.status === 'string' && data.status.trim() ? data.status.trim() : 'unknown';

      let meta: ConversationEntry | undefined;
      try {
        const entries = await listConversations(deps.ctx, deps.participantId);
        meta = entries.find((entry) => entry.id === conversationId);
      } catch {
        meta = undefined;
      }

      let usageText: string;
      try {
        const usage = await deps.ctx.callTool(deps.participantId, 'query_usage', {
          conversationId,
        });
        if (usage.result.status !== 'success') {
          usageText = `Usage: unavailable (${toolFailure(usage.result)})`;
        } else {
          const totals = (usage.result.data as { totals?: Parameters<typeof formatUsage>[0] } | undefined)
            ?.totals;
          usageText = totals ? formatUsage(totals) : 'Usage: no usage recorded yet.';
        }
      } catch (error) {
        usageText = `Usage: unavailable (${error instanceof Error ? error.message : String(error)})`;
      }

      const title = meta?.title ?? fallbackTitle;
      const counts = meta
        ? `Messages: ${meta.messageCount} total, ${activeMessages} active`
        : `Messages: ${activeMessages} active`;
      await deps.send([
        `📊 Context — ${title}`,
        `Conversation: ${conversationId}`,
        `Status: ${status} · ${counts}`,
        usageText,
      ].join('\n'));
      return true;
    }
    default:
      return false;
  }
}
