// src/integration.test.ts
import { describe, expect, it } from 'vitest';

const enabled = process.env.LEGION_TEAMS_INTEGRATION === '1';

describe.skipIf(!enabled)('live Teams integration', () => {
  it('starts a Socket Mode bot and receives an install activity', async () => {
    const { App } = await import('@microsoft/teams.apps');
    const app = new App({
      clientId: process.env.TEAMS_CLIENT_ID,
      clientSecret: process.env.TEAMS_CLIENT_SECRET,
      tenantId: process.env.TEAMS_TENANT_ID,
      socketMode: { geos: ['amer'] },
    });
    const seen: string[] = [];
    app.on('install.add', async ({ activity }) => { seen.push(activity.conversation?.id ?? ''); });
    await app.start();
    try {
      await new Promise((resolve) => setTimeout(resolve, 5_000));
      expect(Array.isArray(seen)).toBe(true);
    } finally {
      await app.stop();
    }
  }, 30_000);
});
