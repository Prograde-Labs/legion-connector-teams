import { describe, expect, it } from 'vitest';
import { connector } from './index.js';
import type { ConnectorRuntimeDeps } from '@legion-collective/core';

const deps = (overrides: Partial<ConnectorRuntimeDeps> = {}): ConnectorRuntimeDeps =>
  ({ collective: {} as never, secrets: {} as never, ...overrides });

describe('connector factory', () => {
  it('returns a connector named teams', () => {
    expect(connector({}, deps()).name).toBe('teams');
  });

  it('throws at start without a secret store', async () => {
    const c = connector({}, deps({ secrets: undefined }));
    await expect(c.start({} as never)).rejects.toThrow(/LEGION_SECRETS_KEY/);
  });

  it('rejects fallback policy without fallbackUserId', async () => {
    const c = connector({ unknownUserPolicy: 'fallback' }, deps());
    await expect(c.start({} as never)).rejects.toThrow(/fallbackUserId/);
  });
});
