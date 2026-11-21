// src/bot-pool.test.ts
import { describe, expect, it, vi } from 'vitest';
import { BotPool, type TeamsAppLike } from './bot-pool.js';
import { resolveOptions } from './options.js';
import type { BotEntry } from './directory.js';
import type { AddressInfo } from 'node:net';

function fakeApp(): TeamsAppLike & { sent: unknown[]; stopped: number; initialized: number } {
  return {
    sent: [],
    stopped: 0,
    initialized: 0,
    async initialize() { this.initialized++; },
    async start() { this.initialized++; },
    async stop() { this.stopped++; },
    async send(_conversationId: string, activity: unknown) { this.sent.push(activity); },
    api: { conversations: { create: async () => ({ id: 'conv' }) } },
    on() {},
  };
}

const bot = (participantId: string, appId: string): BotEntry => ({
  participantId, name: participantId, appId,
  credentials: { appId, clientSecret: 's', tenantId: 't' },
});

const handlers = {
  onMessage: vi.fn(async () => undefined),
  onInstall: vi.fn(async () => undefined),
  onConversationUpdate: vi.fn(async () => undefined),
  onApproval: vi.fn(async () => undefined),
};

describe('BotPool', () => {
  it.each(['construct', 'stop'])('redacts the bot client secret in %s diagnostics', async operation => {
    const entry = bot('assistant', 'app');
    entry.credentials.clientSecret = 'super-secret';
    const fail = () => { throw new Error('clientSecret=super-secret'); };
    const log = vi.spyOn(console, 'error').mockImplementation(() => {});
    const pool = new BotPool(resolveOptions({ transport: 'socket' }), () => {
      if (operation === 'construct') return fail();
      return { ...fakeApp(), stop: fail };
    }, handlers);
    try {
      await pool.start(new Map([['app', entry]]));
      await pool.stop();
      expect(log).toHaveBeenCalledWith(`[teams] bot app ${operation} failed: Error: clientSecret=[redacted]`);
      expect(JSON.stringify(log.mock.calls)).not.toContain('super-secret');
    } finally { log.mockRestore(); }
  });

  it('serves unauthenticated health checks on the shared listener', async () => {
    const pool = new BotPool(resolveOptions({ port: 0 }), fakeApp, handlers);
    await pool.start(new Map([['a', bot('a', 'a')]]));
    try {
      const { port } = pool['server']!.address() as AddressInfo;
      expect((await fetch(`http://127.0.0.1:${port}/health`)).status).toBe(200);
    } finally { await pool.stop(); }
  });

  it.each(['http', 'socket'] as const)('isolates %s bot lifecycle failures and redacts errors', async (transport) => {
    const log = vi.spyOn(console, 'error').mockImplementation(() => {});
    const fail = () => { throw new Error('Bearer sentinel-token'); };
    const healthy = fakeApp();
    const pool = new BotPool(resolveOptions({ port: 0, transport }), ({ appId }) => {
      if (appId === 'factory') return fail();
      if (appId === 'startup') return { ...fakeApp(), initialize: fail, start: fail };
      if (appId === 'shutdown') return { ...fakeApp(), stop: fail };
      return healthy;
    }, handlers);
    try {
      await pool.start(new Map(['factory', 'startup', 'shutdown', 'healthy'].map(id => [id, bot(id, id)])));
      expect(healthy.initialized).toBe(1);
      expect(pool.getByAppId('startup')).toBeUndefined();
      await pool.stop();
      expect(healthy.stopped).toBe(1);
      expect(pool.getByAppId('shutdown')).toBeUndefined();
      expect(log).toHaveBeenCalledTimes(3);
      expect(JSON.stringify(log.mock.calls)).not.toContain('sentinel-token');
    } finally { log.mockRestore(); }
  });

  it('rejects an occupied port and stops initialized bots', async () => {
    const owner = new BotPool(resolveOptions({ port: 0 }), fakeApp, handlers);
    await owner.start(new Map());
    try {
      const { port } = owner['server']!.address() as AddressInfo;
      const app = fakeApp();
      const pool = new BotPool(resolveOptions({ port }), () => app, handlers);
      await expect(pool.start(new Map([['a', bot('a', 'a')]]))).rejects.toMatchObject({ code: 'EADDRINUSE' });
      expect(app.initialized).toBe(1);
      expect(app.stopped).toBe(1);
      expect(pool.getByAppId('a')).toBeUndefined();
    } finally { await owner.stop(); }
  });

  it('starts an app per bot and stops them', async () => {
    const apps: Array<ReturnType<typeof fakeApp>> = [];
    const pool = new BotPool(
      resolveOptions({ stateDir: '/tmp/x', port: 0 }),
      () => { const app = fakeApp(); apps.push(app); return app; },
      handlers,
    );
    await pool.start(new Map([['app', bot('assistant', 'app')]]));
    expect(apps).toHaveLength(1);
    expect(pool.getByAppId('app')).toBe(apps[0]);
    await pool.stop();
    expect(apps[0].stopped).toBe(1);
  });

  it('reconciles socket apps individually', async () => {
    const apps = new Map<string, ReturnType<typeof fakeApp>>();
    const pool = new BotPool(
      resolveOptions({ stateDir: '/tmp/x', transport: 'socket' }),
      (entry) => { const app = fakeApp(); apps.set(entry.appId, app); return app; },
      handlers,
    );
    await pool.start(new Map([['a', bot('a', 'a')]]));
    await pool.reconcile(new Map([['b', bot('b', 'b')]]));
    expect(pool.getByAppId('a')).toBeUndefined();
    expect(pool.getByAppId('b')).toBeDefined();
  });
});
