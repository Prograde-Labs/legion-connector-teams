// src/secrets.test.ts
import { describe, expect, it } from 'vitest';
import { parseBotSecret, resolveBotCredentials } from './secrets.js';

describe('bot secrets', () => {
  it('parses a valid credential JSON', () => {
    expect(parseBotSecret('{"appId":"a","clientSecret":"s","tenantId":"t"}'))
      .toEqual({ appId: 'a', clientSecret: 's', tenantId: 't' });
  });

  it('returns null for missing or malformed values', () => {
    expect(parseBotSecret(null)).toBeNull();
    expect(parseBotSecret('not-json')).toBeNull();
    expect(parseBotSecret('{"appId":"a"}')).toBeNull();
  });

  it.each(['appId', 'clientSecret', 'tenantId'] as const)(
    'rejects empty and whitespace-only %s',
    (field) => {
      for (const blank of ['', ' \t\n']) {
        const credentials = { appId: 'a', clientSecret: 's', tenantId: 't', [field]: blank };
        expect(parseBotSecret(JSON.stringify(credentials))).toBeNull();
      }
    },
  );

  it('reads the participant-scoped secret key', async () => {
    const secrets = {
      getSecret: async (participantId: string, key: string) =>
        participantId === 'assistant' && key === 'teams-bot/app-1'
          ? '{"appId":"app-1","clientSecret":"s","tenantId":"t"}'
          : null,
    };
    await expect(resolveBotCredentials(secrets, 'assistant', 'app-1'))
      .resolves.toEqual({ appId: 'app-1', clientSecret: 's', tenantId: 't' });
    await expect(resolveBotCredentials(secrets, 'assistant', 'app-2')).resolves.toBeNull();
  });
});
