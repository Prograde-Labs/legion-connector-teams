import { describe, expect, it } from 'vitest';
import { consumePendingPair, generatePairCode, pruneExpiredPairs, rememberPendingPair, validatePendingPair } from './pairing.js';
import { emptyState } from './state.js';

describe('pairing codes', () => {
  it.each(['validate', 'consume', 'prune'])('removes malformed null entries during %s', (operation) => {
    const state = emptyState();
    state.pendingPairs['CODE'] = null as never;
    if (operation === 'validate') expect(validatePendingPair(state, 'code', 2000)).toBeNull();
    if (operation === 'consume') expect(consumePendingPair(state, 'code', 2000)).toBeNull();
    if (operation === 'prune') pruneExpiredPairs(state, 2000);
    expect('CODE' in state.pendingPairs).toBe(false);
  });

  it('rejects a non-string expiry', () => {
    const state = emptyState();
    state.pendingPairs['CODE'] = { appId: 'app', connector: 'teams', externalId: 'aad', expiresAt: 123 as never };
    expect(validatePendingPair(state, 'CODE', 2000)).toBeNull();
    expect('CODE' in state.pendingPairs).toBe(false);
  });

  it('validates a code without consuming it, case-insensitively', () => {
    const state = emptyState();
    const code = rememberPendingPair(state, { appId: 'app', connector: 'teams', externalId: 'aad' }, 600, 1000);
    expect(validatePendingPair(state, code.toLowerCase(), 2000)).toEqual(state.pendingPairs[code]);
    expect(state.pendingPairs[code]).toBeDefined();
  });

  it.each(['invalid', '1970-01-01T00:00:01.000Z'])('removes invalid expiry %s during validation', (expiresAt) => {
    const state = emptyState();
    state.pendingPairs['ABC234'] = { appId: 'app', connector: 'teams', externalId: 'aad', expiresAt };
    expect(validatePendingPair(state, 'abc234', 2000)).toBeNull();
    expect(state.pendingPairs['ABC234']).toBeUndefined();
  });

  it('generates six unambiguous characters', () => {
    for (let i = 0; i < 20; i++) {
      expect(generatePairCode(() => 0)).toMatch(/^[2-9A-HJ-NP-Z]{6}$/);
    }
  });

  it('uses injected random values as alphabet indices', () => {
    expect(generatePairCode(() => 1)).toBe('333333');
  });

  it('reuses an unexpired code for the same pair', () => {
    const state = emptyState();
    const first = rememberPendingPair(state, { appId: 'app', connector: 'teams', externalId: 'aad' }, 600, 1000);
    const second = rememberPendingPair(state, { appId: 'app', connector: 'teams', externalId: 'aad' }, 600, 2000);
    expect(second).toBe(first);
  });

  it('replaces a pending code with a malformed expiry', () => {
    const state = emptyState();
    state.pendingPairs['INVALID'] = { appId: 'app', connector: 'teams', externalId: 'aad', expiresAt: 'not-a-date' };
    const code = rememberPendingPair(state, { appId: 'app', connector: 'teams', externalId: 'aad' }, 600, 1000);
    expect(code).not.toBe('INVALID');
    expect(state.pendingPairs['INVALID']).toBeUndefined();
  });

  it('keeps codes separate by connector and external id', () => {
    const state = emptyState();
    const personal = rememberPendingPair(
      state, { appId: 'app', connector: 'teams', externalId: 'aad-1' }, 600, 1000,
    );
    const group = rememberPendingPair(
      state, { appId: 'app', connector: 'teams-group', externalId: 'aad-1' }, 600, 1000,
    );
    expect(group).not.toBe(personal);
    expect(state.pendingPairs[personal]).toMatchObject({ connector: 'teams', externalId: 'aad-1' });
    expect(state.pendingPairs[group]).toMatchObject({ connector: 'teams-group', externalId: 'aad-1' });
  });

  it('consumes a code once', () => {
    const state = emptyState();
    const code = rememberPendingPair(state, { appId: 'app', connector: 'teams', externalId: 'aad' }, 600, 1000);
    expect(consumePendingPair(state, code, 2000)).toMatchObject({ appId: 'app', connector: 'teams', externalId: 'aad' });
    expect(consumePendingPair(state, code, 2000)).toBeNull();
  });

  it('rejects an expired code', () => {
    const state = emptyState();
    const code = rememberPendingPair(state, { appId: 'app', connector: 'teams', externalId: 'aad' }, 10, 1000);
    expect(consumePendingPair(state, code, 12_000)).toBeNull();
  });

  it('rejects and consumes a code with a malformed expiry', () => {
    const state = emptyState();
    state.pendingPairs['ABC234'] = { appId: 'app', connector: 'teams', externalId: 'aad', expiresAt: 'invalid' };
    expect(consumePendingPair(state, 'ABC234', 1000)).toBeNull();
    expect(state.pendingPairs['ABC234']).toBeUndefined();
  });
});
