import { describe, expect, it } from 'vitest';
import { consumePendingPair, generatePairCode, rememberPendingPair } from './pairing.js';
import { emptyState } from './state.js';

describe('pairing codes', () => {
  it('generates six unambiguous characters', () => {
    for (let i = 0; i < 20; i++) {
      expect(generatePairCode(() => 0)).toMatch(/^[2-9A-HJ-NP-Z]{6}$/);
    }
  });

  it('reuses an unexpired code for the same pair', () => {
    const state = emptyState();
    const first = rememberPendingPair(state, 'app', 'aad', 600, 1000);
    const second = rememberPendingPair(state, 'app', 'aad', 600, 2000);
    expect(second).toBe(first);
  });

  it('consumes a code once', () => {
    const state = emptyState();
    const code = rememberPendingPair(state, 'app', 'aad', 600, 1000);
    expect(consumePendingPair(state, code, 2000)).toMatchObject({ appId: 'app', aadObjectId: 'aad' });
    expect(consumePendingPair(state, code, 2000)).toBeNull();
  });

  it('rejects an expired code', () => {
    const state = emptyState();
    const code = rememberPendingPair(state, 'app', 'aad', 10, 1000);
    expect(consumePendingPair(state, code, 12_000)).toBeNull();
  });
});
