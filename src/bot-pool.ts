// src/bot-pool.ts
import { createServer, type Server } from 'node:http';
import express from 'express';
import { App, ExpressAdapter } from '@microsoft/teams.apps';
import type { BotEntry } from './directory.js';
import type { ResolvedOptions } from './options.js';

export interface TeamsAppLike {
  initialize(): Promise<void>;
  start(): Promise<void>;
  stop(): Promise<void>;
  send(conversationId: string, activity: unknown): Promise<unknown>;
  api: { conversations: { create(params: Record<string, unknown>): Promise<{ id: string }> } };
  on(event: string, handler: (ctx: unknown) => Promise<void> | void): void;
}

export interface AppHandlers {
  onMessage: (entry: BotEntry, ctx: never) => Promise<void>;
  onInstall: (entry: BotEntry, ctx: never) => Promise<void>;
  onConversationUpdate: (entry: BotEntry, ctx: never) => Promise<void>;
  onApproval: (entry: BotEntry, ctx: never, decision: 'approve' | 'reject') => Promise<void>;
}

export type TransportContext =
  | { kind: 'http'; expressApp: express.Express }
  | { kind: 'socket'; geos: string[] };

export type TeamsAppFactory = (
  entry: BotEntry,
  handlers: AppHandlers,
  transport: TransportContext,
) => TeamsAppLike;

export function defaultAppFactory(
  entry: BotEntry,
  handlers: AppHandlers,
  transport: TransportContext,
): TeamsAppLike {
  const app = transport.kind === 'http'
    ? new App({
        clientId: entry.credentials.appId,
        clientSecret: entry.credentials.clientSecret,
        tenantId: entry.credentials.tenantId,
        httpServerAdapter: new ExpressAdapter(transport.expressApp),
        messagingEndpoint: `/teams/${entry.participantId}/api/messages`,
      })
    : new App({
        clientId: entry.credentials.appId,
        clientSecret: entry.credentials.clientSecret,
        tenantId: entry.credentials.tenantId,
        socketMode: { geos: transport.geos },
      });

  app.on('message', (ctx) => handlers.onMessage(entry, ctx as never));
  app.on('install.add', (ctx) => handlers.onInstall(entry, ctx as never));
  app.on('conversationUpdate', (ctx) => handlers.onConversationUpdate(entry, ctx as never));
  app.on('card.action.approve', async (ctx) => {
    await handlers.onApproval(entry, ctx as never, 'approve');
    return undefined;
  });
  app.on('card.action.reject', async (ctx) => {
    await handlers.onApproval(entry, ctx as never, 'reject');
    return undefined;
  });

  return app as unknown as TeamsAppLike;
}

export class BotPool {
  private readonly apps = new Map<string, TeamsAppLike>();
  private server?: Server;

  constructor(
    private readonly options: ResolvedOptions,
    private readonly factory: TeamsAppFactory,
    private readonly handlers: AppHandlers,
  ) {}

  async start(bots: Map<string, BotEntry>): Promise<void> {
    if (this.options.transport === 'http') {
      const expressApp = express();
      const server = createServer(expressApp);
      for (const entry of bots.values()) {
        const app = this.factory(entry, this.handlers, { kind: 'http', expressApp });
        await app.initialize();
        this.apps.set(entry.appId, app);
      }
      await new Promise<void>((resolve) => server.listen(this.options.port, '0.0.0.0', resolve));
      this.server = server;
      return;
    }

    for (const entry of bots.values()) {
      const app = this.factory(entry, this.handlers, {
        kind: 'socket',
        geos: this.options.socketGeos,
      });
      await app.start();
      this.apps.set(entry.appId, app);
    }
  }

  async reconcile(bots: Map<string, BotEntry>): Promise<void> {
    if (this.options.transport === 'http') {
      await this.stop();
      await this.start(bots);
      return;
    }

    for (const [appId, app] of [...this.apps]) {
      if (!bots.has(appId)) {
        await app.stop();
        this.apps.delete(appId);
      }
    }
    for (const [appId, entry] of bots) {
      if (this.apps.has(appId)) continue;
      const app = this.factory(entry, this.handlers, {
        kind: 'socket',
        geos: this.options.socketGeos,
      });
      await app.start();
      this.apps.set(appId, app);
    }
  }

  getByAppId(appId: string): TeamsAppLike | undefined {
    return this.apps.get(appId);
  }

  async stop(): Promise<void> {
    for (const app of this.apps.values()) await app.stop();
    this.apps.clear();
    if (this.server) {
      await new Promise<void>((resolve) => this.server!.close(() => resolve()));
      this.server = undefined;
    }
  }
}
