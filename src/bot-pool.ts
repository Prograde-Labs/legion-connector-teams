// src/bot-pool.ts
import { createServer, type Server } from 'node:http';
import express from 'express';
import { App, ExpressAdapter } from '@microsoft/teams.apps';
import type { AdaptiveCardActionMessageResponse } from '@microsoft/teams.api';
import type { BotEntry } from './directory.js';
import { describeError } from './errors.js';
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
  onApproval: (entry: BotEntry, ctx: never, decision: 'approve' | 'reject') => Promise<CardInvokeResponse>;
}

export interface CardInvokeResponse {
  status: 200;
  body: AdaptiveCardActionMessageResponse;
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
  app.on('card.action.approve', (ctx) => handlers.onApproval(entry, ctx as never, 'approve'));
  app.on('card.action.reject', (ctx) => handlers.onApproval(entry, ctx as never, 'reject'));

  return app as unknown as TeamsAppLike;
}

function sameEntry(current: BotEntry, next: BotEntry | undefined): boolean {
  return next !== undefined && current.participantId === next.participantId &&
    current.appId === next.appId && current.credentials.clientSecret === next.credentials.clientSecret &&
    current.credentials.tenantId === next.credentials.tenantId;
}

export class BotPool {
  private readonly apps = new Map<string, { entry: BotEntry; app: TeamsAppLike }>();
  private server?: Server;

  constructor(
    private readonly options: ResolvedOptions,
    private readonly factory: TeamsAppFactory,
    private readonly handlers: AppHandlers,
  ) {}

  private async add(entry: BotEntry, transport: TransportContext): Promise<void> {
    let operation = 'construct';
    let app: TeamsAppLike | undefined;
    try {
      app = this.factory(entry, this.handlers, transport);
      operation = transport.kind === 'http' ? 'initialize' : 'start';
      if (transport.kind === 'http') await app.initialize();
      else await app.start();
      this.apps.set(entry.appId, { entry, app });
    } catch (error) {
      const diagnostic = describeError(error, [entry.credentials.clientSecret]);
      console.error(`[teams] bot ${entry.appId} ${operation} failed: ${diagnostic}`);
      if (app) await this.stopApp(entry, app);
    }
  }

  private async stopApp(entry: BotEntry, app: TeamsAppLike): Promise<void> {
    try {
      await app.stop();
    } catch (error) {
      const diagnostic = describeError(error, [entry.credentials.clientSecret]);
      console.error(`[teams] bot ${entry.appId} stop failed: ${diagnostic}`);
    }
  }

  async start(bots: Map<string, BotEntry>): Promise<void> {
    if (this.options.transport === 'http') {
      const expressApp = express();
      expressApp.get('/health', (_req, res) => res.sendStatus(200));
      const server = createServer(expressApp);
      for (const entry of bots.values()) {
        await this.add(entry, { kind: 'http', expressApp });
      }
      try {
        await new Promise<void>((resolve, reject) => {
          server.once('error', reject);
          server.listen(this.options.port, '0.0.0.0', () => {
            server.off('error', reject);
            resolve();
          });
        });
      } catch (error) {
        await this.stop();
        throw error;
      }
      this.server = server;
      return;
    }

    for (const entry of bots.values()) {
      await this.add(entry, { kind: 'socket', geos: this.options.socketGeos });
    }
  }

  async reconcile(bots: Map<string, BotEntry>): Promise<void> {
    if (this.options.transport === 'http') {
      // Only initialized apps are in this.apps, so a failed bot forces a retry.
      if (this.server && this.apps.size === bots.size &&
          [...this.apps].every(([id, { entry }]) => sameEntry(entry, bots.get(id)))) return;
      await this.stop();
      await this.start(bots);
      return;
    }

    for (const [appId, { entry, app }] of [...this.apps]) {
      if (!sameEntry(entry, bots.get(appId))) {
        this.apps.delete(appId);
        await this.stopApp(entry, app);
      }
    }
    for (const [appId, entry] of bots) {
      if (this.apps.has(appId)) continue;
      await this.add(entry, { kind: 'socket', geos: this.options.socketGeos });
    }
  }

  getByAppId(appId: string): TeamsAppLike | undefined {
    return this.apps.get(appId)?.app;
  }

  ownerOf(appId: string): string | undefined {
    return this.apps.get(appId)?.entry.participantId;
  }

  async stop(): Promise<void> {
    for (const { entry, app } of this.apps.values()) await this.stopApp(entry, app);
    this.apps.clear();
    if (this.server) {
      await new Promise<void>((resolve) => this.server!.close(() => resolve()));
      this.server = undefined;
    }
  }
}
