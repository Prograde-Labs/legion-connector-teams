import { randomInt } from 'node:crypto';
import type { PendingPair, TeamsState } from './state.js';

const ALPHABET = '23456789ABCDEFGHJKLMNPQRSTUVWXYZ';

export function generatePairCode(random: () => number = () => randomInt(0, ALPHABET.length)): string {
  let code = '';
  for (let i = 0; i < 6; i++) code += ALPHABET[random()];
  return code;
}

export function pruneExpiredPairs(state: TeamsState, now = Date.now()): void {
  for (const [code, pair] of Object.entries(state.pendingPairs)) {
    if (Date.parse(pair.expiresAt) <= now) delete state.pendingPairs[code];
  }
}

export function rememberPendingPair(
  state: TeamsState,
  appId: string,
  aadObjectId: string,
  ttlSeconds: number,
  now = Date.now(),
): string {
  pruneExpiredPairs(state, now);
  for (const [code, pair] of Object.entries(state.pendingPairs)) {
    if (pair.appId === appId && pair.aadObjectId === aadObjectId) return code;
  }
  let code = generatePairCode();
  while (state.pendingPairs[code]) code = generatePairCode();
  state.pendingPairs[code] = {
    appId,
    aadObjectId,
    expiresAt: new Date(now + ttlSeconds * 1000).toISOString(),
  };
  return code;
}

export function consumePendingPair(
  state: TeamsState,
  code: string,
  now = Date.now(),
): PendingPair | null {
  const key = code.toUpperCase();
  const pair = state.pendingPairs[key];
  if (!pair) return null;
  delete state.pendingPairs[key];
  const expiry = Date.parse(pair.expiresAt);
  return Number.isNaN(expiry) || expiry <= now ? null : pair;
}
