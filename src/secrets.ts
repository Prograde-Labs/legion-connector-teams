// src/secrets.ts
import type { SecretStoreLike } from './directory.js';

export interface BotCredentials {
  appId: string;
  clientSecret: string;
  tenantId: string;
}

export function parseBotSecret(raw: string | null): BotCredentials | null {
  if (!raw) return null;
  try {
    const value = JSON.parse(raw) as Record<string, unknown>;
    if (
      typeof value.appId !== 'string' ||
      typeof value.clientSecret !== 'string' ||
      typeof value.tenantId !== 'string'
    ) {
      return null;
    }
    return {
      appId: value.appId,
      clientSecret: value.clientSecret,
      tenantId: value.tenantId,
    };
  } catch {
    return null;
  }
}

export async function resolveBotCredentials(
  secrets: SecretStoreLike,
  participantId: string,
  appId: string,
): Promise<BotCredentials | null> {
  return parseBotSecret(await secrets.getSecret(participantId, `teams-bot/${appId}`));
}
