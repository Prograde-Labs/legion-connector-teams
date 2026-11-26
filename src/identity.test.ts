import { describe, expect, it } from 'vitest';
import { findGroupActor, provisionedGroupId, provisionedUserId, resolveGroupIdentity, resolveIdentity } from './identity.js';
import { emptyState } from './state.js';
import { resolveOptions } from './options.js';
import type { CollectiveLike } from '@legion-collective/core';

const collective = (participants: Array<Record<string, unknown>>) =>
  ({
    list: () => participants,
    get: (id: string) => participants.find((p) => p.id === id),
    findByIdentity: (connector: string, externalId: string) =>
      participants.find((p) =>
        (p.identities as Array<{ connector: string; externalId: string }> | undefined)?.some(
          (i) => i.connector === connector && i.externalId === externalId,
        ),
      ),
  }) as unknown as CollectiveLike;

const options = (overrides = {}) => resolveOptions({ stateDir: '/tmp/x', ...overrides });

describe('resolveIdentity', () => {
  it('resolves an explicit teams identity', () => {
    const outcome = resolveIdentity({
      collective: collective([
        { id: 'alice', name: 'Alice', type: 'user', tools: {}, status: 'active',
          identities: [{ connector: 'teams', externalId: 'aad-1' }] },
      ]),
      state: emptyState(), appId: 'app', aadObjectId: 'aad-1', options: options(),
    });
    expect(outcome).toEqual({
      kind: 'resolved',
      user: { participantId: 'alice', aadObjectId: 'aad-1', displayName: undefined },
    });
  });

  it('issues a pairing code for an unknown user', () => {
    const state = emptyState();
    const outcome = resolveIdentity({
      collective: collective([]), state, appId: 'app', aadObjectId: 'aad-2', options: options(),
    });
    expect(outcome.kind).toBe('pairing');
    expect(Object.keys(state.pendingPairs)).toHaveLength(1);
  });

  it('returns auto-provision for that policy', () => {
    const outcome = resolveIdentity({
      collective: collective([]), state: emptyState(), appId: 'app', aadObjectId: 'aad-3',
      displayName: 'New User', options: options({ unknownUserPolicy: 'auto-provision' }),
    });
    expect(outcome).toEqual({ kind: 'auto-provision', aadObjectId: 'aad-3', displayName: 'New User' });
  });

  it('rejects for the reject policy', () => {
    const outcome = resolveIdentity({
      collective: collective([]), state: emptyState(), appId: 'app', aadObjectId: 'aad-4',
      options: options({ unknownUserPolicy: 'reject' }),
    });
    expect(outcome).toEqual({ kind: 'reject', aadObjectId: 'aad-4' });
  });

  it('uses the fallback participant', () => {
    const outcome = resolveIdentity({
      collective: collective([
        { id: 'operator', name: 'Operator', type: 'user', tools: {}, status: 'active' },
      ]),
      state: emptyState(), appId: 'app', aadObjectId: 'aad-5',
      options: options({ unknownUserPolicy: 'fallback', fallbackUserId: 'operator' }),
    });
    expect(outcome).toMatchObject({ kind: 'fallback', user: { participantId: 'operator' } });
  });

  it('treats a retired participant as unknown', () => {
    const outcome = resolveIdentity({
      collective: collective([
        { id: 'old', name: 'Old', type: 'user', tools: {}, status: 'retired',
          identities: [{ connector: 'teams', externalId: 'aad-6' }] },
      ]),
      state: emptyState(), appId: 'app', aadObjectId: 'aad-6',
      options: options({ unknownUserPolicy: 'reject' }),
    });
    expect(outcome.kind).toBe('reject');
  });
});

describe('group identity', () => {
  const groupIdentity = (id: string, conversationId: string) => ({
    id,
    name: 'Engineering',
    type: 'user',
    tools: {},
    status: 'active',
    identities: [{ connector: 'teams-group', externalId: conversationId }],
  });

  it('resolves a registered group by conversation id', () => {
    const outcome = resolveGroupIdentity({
      collective: collective([groupIdentity('team-eng', '19:g@thread.v2')]),
      state: emptyState(),
      appId: 'app',
      conversationId: '19:g@thread.v2',
      name: 'Engineering',
      options: options(),
    });
    expect(outcome).toEqual({
      kind: 'resolved',
      group: { participantId: 'team-eng', conversationId: '19:g@thread.v2', name: 'Engineering' },
    });
  });

  it('issues a group pairing code for an unknown group', () => {
    const state = emptyState();
    const outcome = resolveGroupIdentity({
      collective: collective([]),
      state,
      appId: 'app',
      conversationId: '19:g@thread.v2',
      options: options(),
    });
    expect(outcome.kind).toBe('pairing');
    expect(Object.values(state.pendingPairs)[0]).toMatchObject({
      appId: 'app',
      connector: 'teams-group',
      externalId: '19:g@thread.v2',
    });
  });

  it('returns group auto-provision for that policy', () => {
    const outcome = resolveGroupIdentity({
      collective: collective([]),
      state: emptyState(),
      appId: 'app',
      conversationId: '19:g@thread.v2',
      name: 'Engineering',
      options: options({ unknownUserPolicy: 'auto-provision' }),
    });
    expect(outcome).toEqual({
      kind: 'auto-provision',
      externalId: '19:g@thread.v2',
      name: 'Engineering',
    });
  });

  it('rejects an unknown group for the reject policy', () => {
    const outcome = resolveGroupIdentity({
      collective: collective([]),
      state: emptyState(),
      appId: 'app',
      conversationId: '19:g@thread.v2',
      options: options({ unknownUserPolicy: 'reject' }),
    });
    expect(outcome).toEqual({ kind: 'reject', externalId: '19:g@thread.v2' });
  });

  it('uses the fallback participant for an unknown group', () => {
    const outcome = resolveGroupIdentity({
      collective: collective([
        { id: 'operator', name: 'Operator', type: 'user', tools: {}, status: 'active' },
      ]),
      state: emptyState(),
      appId: 'app',
      conversationId: '19:g@thread.v2',
      name: 'Engineering',
      options: options({ unknownUserPolicy: 'fallback', fallbackUserId: 'operator' }),
    });
    expect(outcome).toEqual({
      kind: 'fallback',
      group: { participantId: 'operator', conversationId: '19:g@thread.v2', name: 'Engineering' },
    });
  });

  it('treats a retired group participant as unknown', () => {
    const outcome = resolveGroupIdentity({
      collective: collective([
        { ...groupIdentity('team-old', '19:g@thread.v2'), status: 'retired' },
      ]),
      state: emptyState(),
      appId: 'app',
      conversationId: '19:g@thread.v2',
      options: options({ unknownUserPolicy: 'reject' }),
    });
    expect(outcome.kind).toBe('reject');
  });

  it('builds stable, readable provisioning ids', () => {
    expect(provisionedGroupId('Engineering', '19:group@thread.v2'))
      .toBe('teams-group-engineering-eqrgnh');
    expect(provisionedGroupId(undefined, '19:group@thread.v2'))
      .toBe('teams-group-group-eqrgnh');
    expect(provisionedUserId('Alice', 'aad-1')).toBe('teams-alice-1j42js');
  });

  it('finds a group actor by conversation id', () => {
    const found = findGroupActor(collective([groupIdentity('team-eng', '19:g@thread.v2')]), '19:g@thread.v2');
    expect(found?.id).toBe('team-eng');
  });
});

import { resolveChatParticipant } from './identity.js';

const participant = (id: string, connector: string, externalId: string, status = 'active') => ({
  id, name: id, type: 'user', status,
  identities: [{ connector, externalId }],
});

function collectiveOf(participants: Array<ReturnType<typeof participant>>): CollectiveLike {
  return {
    list: () => participants,
    get: (id: string) => participants.find((candidate) => candidate.id === id),
    findByIdentity: (connector: string, externalId: string) =>
      participants.find((candidate) =>
        candidate.identities.some((identity) =>
          identity.connector === connector && identity.externalId === externalId)),
  } as unknown as CollectiveLike;
}

describe('resolveChatParticipant', () => {
  const alice = participant('alice', 'teams', 'aad-1');
  const group = participant('team-eng', 'teams-group', '19:g@thread.v2');
  const operator = participant('operator', 'teams', 'aad-op');

  it('resolves a registered personal user', () => {
    expect(resolveChatParticipant({
      collective: collectiveOf([alice, operator]),
      appId: 'app',
      aadObjectId: 'aad-1',
      isGroup: false,
      options: resolveOptions({ stateDir: '/tmp/x' }),
    })).toEqual({ participantId: 'alice', key: 'app:alice' });
  });

  it('resolves a personal fallback with the AAD suffix', () => {
    expect(resolveChatParticipant({
      collective: collectiveOf([alice, operator]),
      appId: 'app',
      aadObjectId: 'aad-9',
      isGroup: false,
      options: resolveOptions({
        stateDir: '/tmp/x',
        unknownUserPolicy: 'fallback',
        fallbackUserId: 'operator',
      }),
    })).toEqual({ participantId: 'operator', key: 'app:operator:aad-9' });
  });

  it('resolves a registered group', () => {
    expect(resolveChatParticipant({
      collective: collectiveOf([alice, group]),
      appId: 'app',
      conversationId: '19:g@thread.v2',
      isGroup: true,
      options: resolveOptions({ stateDir: '/tmp/x' }),
    })).toEqual({ participantId: 'team-eng', key: 'app:team-eng:19:g@thread.v2' });
  });

  it('resolves a group fallback', () => {
    expect(resolveChatParticipant({
      collective: collectiveOf([alice, operator]),
      appId: 'app',
      conversationId: '19:unknown@thread.v2',
      isGroup: true,
      options: resolveOptions({
        stateDir: '/tmp/x',
        unknownUserPolicy: 'fallback',
        fallbackUserId: 'operator',
      }),
    })).toEqual({ participantId: 'operator', key: 'app:operator:19:unknown@thread.v2' });
  });

  it('returns undefined without a fallback', () => {
    expect(resolveChatParticipant({
      collective: collectiveOf([alice]),
      appId: 'app',
      aadObjectId: 'aad-9',
      isGroup: false,
      options: resolveOptions({ stateDir: '/tmp/x' }),
    })).toBeUndefined();
    expect(resolveChatParticipant({
      collective: collectiveOf([alice]),
      appId: 'app',
      conversationId: '19:unknown@thread.v2',
      isGroup: true,
      options: resolveOptions({ stateDir: '/tmp/x' }),
    })).toBeUndefined();
  });

  it('rejects a retired fallback participant', () => {
    expect(resolveChatParticipant({
      collective: collectiveOf([alice, participant('operator', 'teams', 'aad-op', 'retired')]),
      appId: 'app',
      aadObjectId: 'aad-9',
      isGroup: false,
      options: resolveOptions({
        stateDir: '/tmp/x',
        unknownUserPolicy: 'fallback',
        fallbackUserId: 'operator',
      }),
    })).toBeUndefined();
  });
});
