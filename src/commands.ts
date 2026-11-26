// src/commands.ts
import type { ConnectorContext } from '@legion-collective/core';
import { isAwaitingApproval } from './approvals.js';
import { validatePendingPair } from './pairing.js';
import type { ResolvedOptions } from './options.js';
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
    default:
      return false;
  }
}
