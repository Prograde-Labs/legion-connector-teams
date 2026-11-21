// src/directory.ts
import { resolveBotCredentials, type BotCredentials } from './secrets.js';

export interface DirectoryParticipant {
  id: string;
  name?: string;
  type?: string;
  status?: string;
  identities?: Array<{ connector: string; externalId: string }>;
}

export interface CollectiveLike {
  list(): DirectoryParticipant[];
  get(participantId: string): DirectoryParticipant | undefined;
  findByIdentity(connector: string, externalId: string): DirectoryParticipant | undefined;
}

export interface SecretStoreLike {
  getSecret(participantId: string, key: string): Promise<string | null>;
}

export interface BotEntry {
  participantId: string;
  name: string;
  appId: string;
  credentials: BotCredentials;
}

export interface DirectoryResult {
  bots: Map<string, BotEntry>;
  byParticipant: Map<string, BotEntry>;
  warnings: string[];
}

export async function buildDirectory(
  collective: CollectiveLike,
  secrets: SecretStoreLike,
): Promise<DirectoryResult> {
  const bots = new Map<string, BotEntry>();
  const byParticipant = new Map<string, BotEntry>();
  const warnings: string[] = [];

  for (const participant of collective.list()) {
    if ((participant.status ?? 'active') === 'retired') continue;
    const identity = participant.identities?.find((i) => i.connector === 'teams-bot');
    if (!identity) continue;

    const appId = identity.externalId;
    const existing = bots.get(appId);
    if (existing) {
      warnings.push(`duplicate teams-bot appId ${appId}; keeping ${existing.participantId}`);
      continue;
    }

    const credentials = await resolveBotCredentials(secrets, participant.id, appId);
    if (!credentials) {
      warnings.push(`missing or malformed teams-bot secret for ${participant.id} (${appId})`);
      continue;
    }
    if (credentials.appId !== appId) {
      warnings.push(`teams-bot secret appId mismatch for ${participant.id} (${appId})`);
      continue;
    }

    const entry: BotEntry = {
      participantId: participant.id,
      name: participant.name ?? participant.id,
      appId,
      credentials,
    };
    bots.set(appId, entry);
    byParticipant.set(participant.id, entry);
  }

  return { bots, byParticipant, warnings };
}
