// src/directory.test.ts
import { describe, expect, it } from 'vitest';
import { buildDirectory } from './directory.js';
import type { CollectiveLike, SecretStoreLike } from './directory.js';

const collective = (participants: Array<Record<string, unknown>>) =>
  ({ list: () => participants }) as unknown as CollectiveLike;

const secrets = (values: Record<string, string>): SecretStoreLike =>
  ({ getSecret: async (id: string, key: string) => values[`${id}:${key}`] ?? null }) as SecretStoreLike;

describe('buildDirectory', () => {
  it('includes only participants with resolvable bot credentials', async () => {
    const result = await buildDirectory(
      collective([
        { id: 'assistant', name: 'Assistant', type: 'agent', tools: {}, status: 'active',
          identities: [{ connector: 'teams-bot', externalId: 'app-1' }] },
        { id: 'researcher', name: 'Researcher', type: 'agent', tools: {}, status: 'active',
          identities: [{ connector: 'teams-bot', externalId: 'app-2' }] },
        { id: 'operator', name: 'Operator', type: 'user', tools: {}, status: 'active',
          identities: [{ connector: 'teams', externalId: 'aad-op' }] },
        { id: 'retired', name: 'Retired', type: 'agent', tools: {}, status: 'retired',
          identities: [{ connector: 'teams-bot', externalId: 'app-3' }] },
      ]),
      secrets({ 'assistant:teams-bot/app-1': '{"appId":"app-1","clientSecret":"s","tenantId":"t"}' }),
    );
    expect([...result.bots.keys()]).toEqual(['app-1']);
    expect(result.byParticipant.get('assistant')?.name).toBe('Assistant');
    expect(result.warnings.join(' ')).toMatch(/app-2/);
  });

  it('keeps the first bot on duplicate appId', async () => {
    const result = await buildDirectory(
      collective([
        { id: 'a', name: 'A', type: 'agent', tools: {},
          identities: [{ connector: 'teams-bot', externalId: 'dup' }] },
        { id: 'b', name: 'B', type: 'agent', tools: {},
          identities: [{ connector: 'teams-bot', externalId: 'dup' }] },
      ]),
      secrets({
        'a:teams-bot/dup': '{"appId":"dup","clientSecret":"s","tenantId":"t"}',
        'b:teams-bot/dup': '{"appId":"dup","clientSecret":"s","tenantId":"t"}',
      }),
    );
    expect(result.bots.get('dup')?.participantId).toBe('a');
  });

  it('warns and skips malformed and mismatched secrets', async () => {
    const result = await buildDirectory(
      collective([
        { id: 'a', identities: [{ connector: 'teams-bot', externalId: 'app-a' }] },
        { id: 'b', identities: [{ connector: 'teams-bot', externalId: 'app-b' }] },
      ]),
      secrets({
        'a:teams-bot/app-a': 'not-json',
        'b:teams-bot/app-b': '{"appId":"other","clientSecret":"s","tenantId":"t"}',
      }),
    );
    expect(result.bots.size).toBe(0);
    expect(result.warnings.join(' ')).toMatch(/malformed.*app-a/);
    expect(result.warnings.join(' ')).toMatch(/mismatch.*app-b/);
  });

  it('uses the participant ID when the name is absent', async () => {
    const result = await buildDirectory(
      collective([{ id: 'a', identities: [{ connector: 'teams-bot', externalId: 'app-a' }] }]),
      secrets({ 'a:teams-bot/app-a': '{"appId":"app-a","clientSecret":"s","tenantId":"t"}' }),
    );
    expect(result.byParticipant.get('a')?.name).toBe('a');
  });
});
