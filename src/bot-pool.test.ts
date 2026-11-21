// src/bot-pool.test.ts
import { describe, expect, it, vi } from 'vitest';
import { BotPool, type TeamsAppLike } from './bot-pool.js';
import { resolveOptions } from './options.js';
import type { BotEntry } from './directory.js';

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
