// src/bot-pool.ts
import { createServer, type Server } from 'node:http';
import express from 'express';
import { App, ExpressAdapter, type HttpMethod, type HttpRouteHandler } from '@microsoft/teams.apps';
import type {
  AdaptiveCardActionCardResponse,
  AdaptiveCardActionMessageResponse,
} from '@microsoft/teams.api';
import type { BotEntry } from './directory.js';
import { describeError } from './errors.js';
import type { ResolvedOptions } from './options.js';

export interface TeamsAppLike {
  initialize(): Promise<void>;
  start(): Promise<void>;
  stop(): Promise<void>;
  send(conversationId: string, activity: unknown): Promise<{ id?: string }>;
  api: {
    conversations: {
      create(params: Record<string, unknown>): Promise<{ id: string }>;
      updateActivity(conversationId: string, id: string, params: unknown): Promise<unknown>;
    };
    reactions: {
      add(conversationId: string, activityId: string, type: string): Promise<void>;
      delete(conversationId: string, activityId: string, type: string): Promise<void>;
    };
  };
  on(event: string, handler: (ctx: unknown) => Promise<void> | void): void;
}

export interface AppHandlers {
  onMessage: (entry: BotEntry, ctx: never) => Promise<void>;
  onInstall: (entry: BotEntry, ctx: never) => Promise<void>;
  onConversationUpdate: (entry: BotEntry, ctx: never) => Promise<void>;
  onApproval: (entry: BotEntry, ctx: never, decision: 'approve' | 'reject') => Promise<CardInvokeResponse>;
  onToolAction: (entry: BotEntry, ctx: never) => Promise<CardInvokeResponse>;
}

export interface CardInvokeResponse {
  status: 200;
  body: AdaptiveCardActionMessageResponse | AdaptiveCardActionCardResponse;
}

export type TransportContext =
  | { kind: 'http'; expressApp: express.Express }
  | { kind: 'socket'; geos: string[] };

export type TeamsAppFactory = (
  entry: BotEntry,
  handlers: AppHandlers,
  transport: TransportContext,
) => TeamsAppLike;

const httpRoutes = new WeakMap<express.Express, Set<express.Router>>();

function routesOn(sharedApp: express.Express): Set<express.Router> {
  const existing = httpRoutes.get(sharedApp);
  if (existing) return existing;

  const routes = new Set<express.Router>();
  sharedApp.use((req, res, next) => {
    const remaining = routes.values();
    function dispatch(error?: unknown): void {
      if (error) return next(error);
      const router = remaining.next().value;
      if (router) router(req, res, dispatch);
      else next();
    }
    dispatch();
  });
  httpRoutes.set(sharedApp, routes);
  return routes;
}

// The SDK's stock adapter cannot unregister routes or stop an externally managed
// server. Keep each bot's routes detachable without touching the shared listener.
class BotExpressAdapter extends ExpressAdapter {
  private readonly activeRoutes: Set<express.Router>;

  constructor(sharedApp: express.Express) {
    super(express());
    this.activeRoutes = routesOn(sharedApp);
  }

  override registerRoute(method: HttpMethod, path: string, handler: HttpRouteHandler): void {
    super.registerRoute(method, path, handler);
    this.activeRoutes.add(this.express.router);
  }

  override async stop(): Promise<void> {
    this.activeRoutes.delete(this.express.router);
  }
}

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
        httpServerAdapter: new BotExpressAdapter(transport.expressApp),
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
  app.on('card.action.tool', (ctx) => handlers.onToolAction(entry, ctx as never));

  return app as unknown as TeamsAppLike;
}

function sameEntry(current: BotEntry, next: BotEntry | undefined): boolean {
  return next !== undefined && current.participantId === next.participantId &&
    current.appId === next.appId && current.credentials.clientSecret === next.credentials.clientSecret &&
    current.credentials.tenantId === next.credentials.tenantId;
}

export class BotPool {
  private readonly apps = new Map<string, { entry: BotEntry; app: TeamsAppLike }>();
  private readonly attempted = new Map<string, BotEntry>();
  private expressApp?: express.Express;
  private server?: Server;

  constructor(
    private readonly options: ResolvedOptions,
    private readonly factory: TeamsAppFactory,
    private readonly handlers: AppHandlers,
  ) {}

  private async add(entry: BotEntry, transport: TransportContext): Promise<void> {
    this.attempted.set(entry.appId, entry);
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
    await this.reconcile(bots);
  }

  private transportContext(): TransportContext {
    if (this.options.transport === 'socket') {
      return { kind: 'socket', geos: this.options.socketGeos };
    }
    if (!this.expressApp) {
      this.expressApp = express();
      this.expressApp.get('/health', (_req, res) => res.sendStatus(200));
    }
    return { kind: 'http', expressApp: this.expressApp };
  }

  private async listen(expressApp: express.Express): Promise<void> {
    if (this.server?.listening) return;
    const server = this.server ?? createServer(expressApp);
    try {
      await new Promise<void>((resolve, reject) => {
        server.once('error', reject);
        server.listen(this.options.port, '0.0.0.0', () => {
          server.off('error', reject);
          resolve();
        });
      });
    } catch (error) {
      if (!this.server) await this.stop();
      throw error;
    }
    this.server = server;
  }

  async reconcile(bots: Map<string, BotEntry>): Promise<void> {
    const staleAttempts = [...this.attempted].filter(([appId, entry]) => !sameEntry(entry, bots.get(appId)));
    if (this.options.transport === 'http' && staleAttempts.length > 0) {
      // Injected factories can mount routes directly, even before throwing during construction.
      await this.stop();
      return this.start(bots);
    }

    const transport = this.transportContext();
    for (const [appId, entry] of staleAttempts) {
      this.attempted.delete(appId);
      const app = this.apps.get(appId)?.app;
      this.apps.delete(appId);
      if (app) await this.stopApp(entry, app);
    }
    for (const [appId, entry] of bots) {
      if (this.apps.has(appId)) continue;
      await this.add(entry, transport);
    }
    if (transport.kind === 'http') await this.listen(transport.expressApp);
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
    this.attempted.clear();
    if (this.server) {
      await new Promise<void>((resolve) => this.server!.close(() => resolve()));
      this.server = undefined;
    }
    this.expressApp = undefined;
  }
}
