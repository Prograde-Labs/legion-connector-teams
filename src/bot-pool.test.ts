// src/bot-pool.test.ts
import { describe, expect, it, vi } from 'vitest';
import { BotPool, defaultAppFactory, type AppHandlers, type TeamsAppLike } from './bot-pool.js';
import { App } from '@microsoft/teams.apps';
import express from 'express';
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

const handlers: AppHandlers = {
  onMessage: vi.fn(async () => undefined),
  onInstall: vi.fn(async () => undefined),
  onConversationUpdate: vi.fn(async () => undefined),
  onApproval: async () => ({ status: 200, body: {
    statusCode: 200, type: 'application/vnd.microsoft.activity.message', value: 'Approval recorded.',
  } }),
};

describe('BotPool', () => {
  it('returns the invoke response from the SDK approve and reject event callbacks', async () => {
    const on = vi.spyOn(App.prototype, 'on');
    const entry = bot('assistant', 'app');
    const onApproval = vi.fn(handlers.onApproval);
    try {
      defaultAppFactory(entry, { ...handlers, onApproval }, { kind: 'http', expressApp: express() });
      for (const decision of ['approve', 'reject'] as const) {
        const callback = on.mock.calls.find(([name]) => name === `card.action.${decision}`)?.[1];
        expect(callback).toBeDefined();
        const context = { activity: {} } as never;
        expect(await callback!(context)).toEqual({ status: 200, body: {
          statusCode: 200, type: 'application/vnd.microsoft.activity.message', value: 'Approval recorded.',
        } });
        expect(onApproval).toHaveBeenLastCalledWith(entry, context, decision);
      }
    } finally { on.mockRestore(); }
  });
  it.each(['http', 'socket'] as const)('keeps unchanged %s apps and replaces changed credentials', async (transport) => {
    const factory = vi.fn(fakeApp);
    const pool = new BotPool(resolveOptions({ transport, port: 0 }), factory, handlers);
    await pool.start(new Map([['app', bot('assistant', 'app')]]));
    try {
      const first = pool.getByAppId('app');
      const entry = bot('assistant', 'app');
      await pool.reconcile(new Map([['app', entry]]));
      expect(factory).toHaveBeenCalledTimes(1);
      expect(pool.getByAppId('app')).toBe(first);
      await pool.reconcile(new Map([['app', { ...entry, credentials: { ...entry.credentials, clientSecret: 'rotated' } }]]));
      expect(factory).toHaveBeenCalledTimes(2);
      expect(factory.mock.results[0].value.stopped).toBe(1);
      expect(pool.getByAppId('app')).not.toBe(first);
      expect(pool.ownerOf('app')).toBe('assistant');
    } finally { await pool.stop(); }
  });
  it.each(['construct', 'initialize'])('retries an unchanged HTTP bot after a %s failure', async (operation) => {
    const log = vi.spyOn(console, 'error').mockImplementation(() => {});
    const fail = () => { throw new Error('bot unavailable'); };
    const a = fakeApp();
    const b = fakeApp();
    const createB = vi.fn(() => b).mockImplementationOnce(() =>
      operation === 'construct' ? fail() : { ...fakeApp(), initialize: fail });
    const factory = vi.fn(({ appId }: BotEntry) => appId === 'a' ? a : createB());
    const pool = new BotPool(resolveOptions({ transport: 'http', port: 0 }), factory, handlers);
    const directory = () => new Map([['a', bot('a', 'a')], ['b', bot('b', 'b')]]);
    try {
      await pool.start(new Map([['a', bot('a', 'a')]]));
      const server = pool['server']!;
      const { port } = server.address() as AddressInfo;
      const health = `http://127.0.0.1:${port}/health`;
      expect((await fetch(health)).status).toBe(200);

      await pool.reconcile(directory());
      expect(createB).toHaveBeenCalledTimes(1);
      expect(pool.getByAppId('b')).toBeUndefined();
      expect(pool['server']).toBe(server);
      expect((await fetch(health)).status).toBe(200);

      await pool.reconcile(directory());
      expect(createB).toHaveBeenCalledTimes(2);
      expect(pool.getByAppId('b')).toBe(b);
      expect(b.initialized).toBe(1);
      expect(pool.getByAppId('a')).toBe(a);
      expect(a.initialized).toBe(1);
      expect(a.stopped).toBe(0);
      expect(pool['server']).toBe(server);
      expect((await fetch(health)).status).toBe(200);
    } finally { await pool.stop(); log.mockRestore(); }
  });

  it('stops only the removed HTTP bot', async () => {
    const a = fakeApp();
    const b = fakeApp();
    const pool = new BotPool(resolveOptions({ port: 0 }), ({ appId }) => appId === 'a' ? a : b, handlers);
    try {
      await pool.start(new Map([['a', bot('a', 'a')], ['b', bot('b', 'b')]]));
      const server = pool['server'];

      await pool.reconcile(new Map([['a', bot('a', 'a')]]));

      expect(b.stopped).toBe(1);
      expect(pool.getByAppId('b')).toBeUndefined();
      expect(a.stopped).toBe(0);
      expect(a.initialized).toBe(1);
      expect(pool.getByAppId('a')).toBe(a);
      expect(pool['server']).toBe(server);
    } finally { await pool.stop(); }
  });

  it('removes the old SDK route when an HTTP bot changes participant', async () => {
    const pool = new BotPool(resolveOptions({ port: 0 }), defaultAppFactory, handlers);
    try {
      await pool.start(new Map([['app', bot('old', 'app')]]));
      const { port } = pool['server']!.address() as AddressInfo;
      const send = (participant: string) => fetch(`http://127.0.0.1:${port}/teams/${participant}/api/messages`, {
        method: 'POST', headers: { 'content-type': 'application/json' }, body: '{}',
      });
      expect((await send('old')).status).toBe(401);

      await pool.reconcile(new Map([['app', bot('new', 'app')]]));

      expect(pool.ownerOf('app')).toBe('new');
      expect((await send('old')).status).toBe(404);
      expect((await send('new')).status).toBe(401);
    } finally { await pool.stop(); }
  });

  it('does nothing when the healthy HTTP directory is unchanged', async () => {
    const factory = vi.fn(fakeApp);
    const pool = new BotPool(resolveOptions({ port: 0 }), factory, handlers);
    const directory = () => new Map([['a', bot('a', 'a')], ['b', bot('b', 'b')]]);
    try {
      await pool.start(directory());
      const server = pool['server'];

      await pool.reconcile(directory());

      expect(factory).toHaveBeenCalledTimes(2);
      for (const [index, id] of ['a', 'b'].entries()) {
        const app = factory.mock.results[index].value;
        expect(pool.getByAppId(id)).toBe(app);
        expect(app.initialized).toBe(1);
        expect(app.stopped).toBe(0);
      }
      expect(pool['server']).toBe(server);
    } finally { await pool.stop(); }
  });

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
