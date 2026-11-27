import type { CollectiveLike } from './directory.js';
import type { ResolvedOptions } from './options.js';
import { rememberPendingPair } from './pairing.js';
import type { TeamsState } from './state.js';

export interface ResolvedUser {
  participantId: string;
  aadObjectId: string;
  displayName?: string;
}

export type IdentityOutcome =
  | { kind: 'resolved'; user: ResolvedUser }
  | { kind: 'pairing'; aadObjectId: string; code: string }
  | { kind: 'reject'; aadObjectId: string }
  | { kind: 'auto-provision'; aadObjectId: string; displayName?: string }
  | { kind: 'fallback'; user: ResolvedUser };

export function readTeamsId(from: { aadObjectId?: string; id?: string } | undefined): string | undefined {
  return from?.aadObjectId?.trim() || from?.id?.trim();
}

export function findTeamsActor(collective: CollectiveLike, aadObjectId: string) {
  const participant = collective.findByIdentity('teams', aadObjectId);
  return participant?.status === 'retired' ? undefined : participant;
}

export interface ResolvedGroup {
  participantId: string;
  conversationId: string;
  name?: string;
}

export type GroupIdentityOutcome =
  | { kind: 'resolved'; group: ResolvedGroup }
  | { kind: 'pairing'; externalId: string; code: string }
  | { kind: 'reject'; externalId: string }
  | { kind: 'auto-provision'; externalId: string; name?: string }
  | { kind: 'fallback'; group: ResolvedGroup };

function slug(value: string): string {
  return value.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '').slice(0, 24) || 'user';
}

export function hash6(value: string): string {
  let hash = 0;
  for (const char of value) hash = (hash * 31 + char.charCodeAt(0)) >>> 0;
  return hash.toString(36).padStart(6, '0').slice(0, 6);
}

export function provisionedUserId(name: string, aadObjectId: string): string {
  return `teams-${slug(name)}-${hash6(aadObjectId)}`;
}

export function provisionedGroupId(name: string | undefined, conversationId: string): string {
  const base = name?.trim() ? name : 'group';
  return `teams-group-${slug(base)}-${hash6(conversationId)}`;
}

export function findGroupActor(collective: CollectiveLike, conversationId: string) {
  const participant = collective.findByIdentity('teams-group', conversationId);
  return participant?.status === 'retired' ? undefined : participant;
}

export function resolveGroupIdentity(args: {
  collective: CollectiveLike;
  state: TeamsState;
  appId: string;
  conversationId: string;
  name?: string;
  options: ResolvedOptions;
  now?: number;
}): GroupIdentityOutcome {
  const { collective, state, appId, conversationId, name, options, now } = args;

  const found = findGroupActor(collective, conversationId);
  if (found) {
    return { kind: 'resolved', group: { participantId: found.id, conversationId, name } };
  }

  if (options.unknownUserPolicy === 'fallback' && options.fallbackUserId) {
    const fallback = collective.get(options.fallbackUserId);
    if (fallback && (fallback.status ?? 'active') !== 'retired') {
      return { kind: 'fallback', group: { participantId: fallback.id, conversationId, name } };
    }
    return { kind: 'reject', externalId: conversationId };
  }

  switch (options.unknownUserPolicy) {
    case 'pairing':
      return {
        kind: 'pairing',
        externalId: conversationId,
        code: rememberPendingPair(
          state,
          { appId, connector: 'teams-group', externalId: conversationId },
          options.pairingTtlSeconds,
          now,
        ),
      };
    case 'auto-provision':
      return { kind: 'auto-provision', externalId: conversationId, name };
    case 'reject':
    case 'fallback':
      return { kind: 'reject', externalId: conversationId };
  }
}

export function resolveIdentity(args: {
  collective: CollectiveLike;
  state: TeamsState;
  appId: string;
  aadObjectId: string;
  displayName?: string;
  options: ResolvedOptions;
  now?: number;
}): IdentityOutcome {
  const { collective, state, appId, aadObjectId, displayName, options, now } = args;

  const found = findTeamsActor(collective, aadObjectId);
  if (found) {
    return {
      kind: 'resolved',
      user: { participantId: found.id, aadObjectId, displayName },
    };
  }

  if (options.unknownUserPolicy === 'fallback' && options.fallbackUserId) {
    const fallback = collective.get(options.fallbackUserId);
    if (fallback && (fallback.status ?? 'active') !== 'retired') {
      return {
        kind: 'fallback',
        user: { participantId: fallback.id, aadObjectId, displayName },
      };
    }
    return { kind: 'reject', aadObjectId };
  }

  switch (options.unknownUserPolicy) {
    case 'pairing':
      return {
        kind: 'pairing',
        aadObjectId,
        code: rememberPendingPair(
          state,
          { appId, connector: 'teams', externalId: aadObjectId },
          options.pairingTtlSeconds,
          now,
        ),
      };
    case 'auto-provision':
      return { kind: 'auto-provision', aadObjectId, displayName };
    case 'reject':
    case 'fallback':
      return { kind: 'reject', aadObjectId };
  }
}

export interface ChatParticipant {
  participantId: string;
  key: string;
}

function fallbackParticipant(
  collective: CollectiveLike,
  options: ResolvedOptions,
) {
  if (options.unknownUserPolicy !== 'fallback' || !options.fallbackUserId) return undefined;
  const fallback = collective.get(options.fallbackUserId);
  return fallback && (fallback.status ?? 'active') !== 'retired' ? fallback : undefined;
}

export function resolveChatParticipant(args: {
  collective: CollectiveLike;
  appId: string;
  aadObjectId?: string;
  conversationId?: string;
  isGroup: boolean;
  options: ResolvedOptions;
}): ChatParticipant | undefined {
  const { collective, appId, aadObjectId, conversationId, isGroup, options } = args;

  if (isGroup) {
    if (!conversationId) return undefined;
    const group = findGroupActor(collective, conversationId);
    if (group) return { participantId: group.id, key: `${appId}:${group.id}:${conversationId}` };
    const fallback = fallbackParticipant(collective, options);
    return fallback
      ? { participantId: fallback.id, key: `${appId}:${fallback.id}:${conversationId}` }
      : undefined;
  }

  const user = aadObjectId ? findTeamsActor(collective, aadObjectId) : undefined;
  if (user) return { participantId: user.id, key: `${appId}:${user.id}` };
  const fallback = fallbackParticipant(collective, options);
  return fallback && aadObjectId
    ? { participantId: fallback.id, key: `${appId}:${fallback.id}:${aadObjectId}` }
    : undefined;
}
