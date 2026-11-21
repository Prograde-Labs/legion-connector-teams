// src/commands.ts
import type { ConnectorContext } from '@legion-collective/core';
import { validatePendingPair } from './pairing.js';
import type { ResolvedOptions } from './options.js';
import type { TeamsState } from './state.js';

export interface CommandDeps {
  appId: string;
  participantId: string;
  aadObjectId: string;
  text: string;
  state: TeamsState;
  ctx: ConnectorContext;
  options: ResolvedOptions;
  send: (text: string) => Promise<void>;
}

export async function handleCommand(deps: CommandDeps): Promise<boolean> {
  const [raw, ...rest] = deps.text.trim().split(/\s+/);
  const command = raw.toLowerCase();

  switch (command) {
    case '/new': {
      delete deps.state.legionConversations[`${deps.appId}:${deps.participantId}`];
      await deps.send('Started a new conversation.');
      return true;
    }
    case '/pair': {
      const code = rest[0]?.toUpperCase();
      const target = rest[1];
      if (!code || !target) {
        await deps.send('Usage: /pair <code> <participant-id>');
        return true;
      }
      const pending = validatePendingPair(deps.state, code);
      if (!pending || pending.appId !== deps.appId || pending.aadObjectId !== deps.aadObjectId) {
        await deps.send('Pairing code is invalid or expired.');
        return true;
      }
      const caller = deps.options.pairingParticipantId ?? deps.participantId;
      const { result } = await deps.ctx.callTool(caller, 'pair', {
        participantId: target,
        connector: 'teams',
        externalId: deps.aadObjectId,
      });
      if (result.status === 'success') {
        delete deps.state.pendingPairs[code.toUpperCase()];
        await deps.send(`Paired ${target} with your Teams identity.`);
      } else if (result.status === 'pending_approval') {
        await deps.send('Pairing is awaiting approval.');
      } else {
        await deps.send(`Pairing failed: ${result.error ?? result.status}`);
      }
      return true;
    }
    case '/whoami': {
      const conversation =
        deps.state.legionConversations[`${deps.appId}:${deps.participantId}`] ?? '(none)';
      await deps.send(
        `Participant: ${deps.participantId}\nTeams ID: ${deps.aadObjectId}\n` +
        `Bot app: ${deps.appId}\nConversation: ${conversation}`,
      );
      return true;
    }
    case '/help': {
      await deps.send(
        'Commands:\n' +
        '/new — start a new conversation\n' +
        '/pair <code> <participant-id> — pair your Teams identity\n' +
        '/whoami — show your identity\n' +
        '/help — this message',
      );
      return true;
    }
    default:
      return false;
  }
}
