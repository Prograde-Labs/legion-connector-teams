// src/integration.test.ts
import { describe, expect, it } from 'vitest';

const enabled = process.env.LEGION_TEAMS_INTEGRATION === '1';

describe.skipIf(!enabled)('live Teams integration', () => {
  it('starts a Socket Mode bot with a ready geo', async () => {
    const { App } = await import('@microsoft/teams.apps');
    const app = new App({
      clientId: process.env.TEAMS_CLIENT_ID,
      clientSecret: process.env.TEAMS_CLIENT_SECRET,
      tenantId: process.env.TEAMS_TENANT_ID,
      socketMode: { geos: ['amer'] },
    });
    const seen: string[] = [];
    app.on('install.add', async ({ activity }) => { seen.push(activity.conversation?.id ?? ''); });
    let interval: ReturnType<typeof setInterval> | undefined;
    let timeout: ReturnType<typeof setTimeout> | undefined;
    try {
      await app.start();
      const socket = app.socketMode;
      expect(socket).toBeDefined();
      if (!socket) throw new Error('Socket Mode adapter is unavailable');
      if (socket.status !== 'ready') {
        await new Promise<void>((resolve, reject) => {
          interval = setInterval(() => {
            if (socket.status === 'ready') resolve();
          }, 100);
          timeout = setTimeout(() => reject(new Error('Socket Mode did not become ready within 20 seconds')), 20_000);
        });
      }
      expect(socket.geoStatuses.some(({ status }) => status === 'ready')).toBe(true);
    } finally {
      clearInterval(interval);
      clearTimeout(timeout);
      await app.stop();
    }
  }, 30_000);
});
