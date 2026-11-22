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
        code: rememberPendingPair(state, appId, aadObjectId, options.pairingTtlSeconds, now),
      };
    case 'auto-provision':
      return { kind: 'auto-provision', aadObjectId, displayName };
    case 'reject':
    case 'fallback':
      return { kind: 'reject', aadObjectId };
  }
}
