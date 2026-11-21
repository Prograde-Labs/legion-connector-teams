// src/connector.ts
import type { Connector, ConnectorContext, ConnectorRuntimeDeps } from '@legion-collective/core';
import { handleApprovalAction } from './approvals.js';
import { BotPool, defaultAppFactory, type TeamsAppFactory } from './bot-pool.js';
import { buildDirectory, type BotEntry } from './directory.js';
import { describeError } from './errors.js';
import { handleInbound, type TeamsContextLike } from './inbound.js';
import { deliverMessage } from './outbound.js';
import { resolveOptions, type ResolvedOptions, type TeamsConnectorOptions } from './options.js';
import { StateStore, type TeamsState } from './state.js';

export class TeamsConnector implements Connector {
  readonly name = 'teams';
  private ctx?: ConnectorContext;
  private store?: StateStore;
  private state?: TeamsState;
  private pool?: BotPool;
  private rescan?: NodeJS.Timeout;
  private stopping = false;
  private inFlight?: Promise<void>;
  private readonly options: ResolvedOptions;

  constructor(
    options: TeamsConnectorOptions,
    private readonly deps: ConnectorRuntimeDeps,
    private readonly appFactory: TeamsAppFactory = defaultAppFactory,
  ) {
    this.options = resolveOptions(options);
  }

  async start(ctx: ConnectorContext): Promise<void> {
    if (!this.deps.secrets) {
      throw new Error('Teams connector requires LEGION_SECRETS_KEY (deps.secrets)');
    }
    if (this.options.unknownUserPolicy === 'fallback' && !this.options.fallbackUserId) {
      throw new Error('unknownUserPolicy "fallback" requires fallbackUserId');
    }

    this.stopping = false;
    this.ctx = ctx;
    this.store = new StateStore(this.options.stateDir);
    this.state = await this.store.load();

    const directory = await buildDirectory(this.deps.collective, this.deps.secrets);
    for (const warning of directory.warnings) console.warn(`[teams] ${warning}`);

    this.pool = new BotPool(this.options, this.appFactory, {
      onMessage: (bot, context) => this.onMessage(bot, context),
      onInstall: (bot, context) => this.onInstall(bot, context),
      onConversationUpdate: (bot, context) => this.onInstall(bot, context),
      onApproval: (bot, context, decision) => this.onApproval(bot, context, decision),
    });
    await this.pool.start(directory.bots);

    this.rescan = setInterval(() => void this.rebuild(), this.options.rescanIntervalMs);
    this.rescan.unref?.();
  }

  private async rebuild(): Promise<void> {
    if (this.stopping || this.inFlight || !this.deps.secrets) return;
    this.inFlight = buildDirectory(this.deps.collective, this.deps.secrets)
      .then(async (directory) => {
        if (this.stopping) return;
        for (const warning of directory.warnings) console.warn(`[teams] ${warning}`);
        await this.pool?.reconcile(directory.bots);
      })
      .catch((error) => console.error(`[teams] rebuild failed: ${describeError(error)}`))
      .finally(() => { this.inFlight = undefined; });
    await this.inFlight;
  }

  private async onMessage(bot: BotEntry, context: TeamsContextLike): Promise<void> {
    if (!this.ctx || !this.state || !this.store) return;
    await handleInbound(
      {
        bot,
        ctx: this.ctx,
        collective: this.deps.collective,
        state: this.state,
        stateStore: this.store,
        options: this.options,
      },
      context,
    );
  }

  private async onInstall(bot: BotEntry, context: TeamsContextLike): Promise<void> {
    if (!this.state || !this.store) return;
    const activity = context.activity;
    if (activity.conversation?.conversationType !== 'personal' || !activity.conversation.id?.trim()) return;
    const aadObjectId = activity.from?.aadObjectId ?? activity.from?.id;
    if (!aadObjectId) return;
    const participant = this.deps.collective.findByIdentity('teams', aadObjectId);
    if (!participant) return;
    this.state.conversations[`${bot.appId}:${participant.id}`] = {
      conversationId: activity.conversation.id,
      tenantId: activity.channelData?.tenant?.id,
      serviceUrl: activity.serviceUrl,
      aadObjectId,
      displayName: activity.from?.name,
      updatedAt: new Date().toISOString(),
    };
    await this.store.save(this.state);
  }

  private async onApproval(
    _bot: BotEntry,
    context: TeamsContextLike,
    decision: 'approve' | 'reject',
  ): Promise<void> {
    if (!this.ctx) return;
    const aadObjectId = context.activity.from?.aadObjectId ?? context.activity.from?.id;
    if (!aadObjectId) return;
    const participant = this.deps.collective.findByIdentity('teams', aadObjectId);
    if (!participant) {
      await context.send('Not registered.');
      return;
    }
    const data = context.activity.value?.action?.data as { approvalId?: string } | undefined;
    if (!data?.approvalId) {
      await context.send('Missing approval id.');
      return;
    }
    const result = await handleApprovalAction({
      ctx: this.ctx,
      userId: participant.id,
      decision,
      approvalId: data.approvalId,
    });
    await context.send(
      result.status === 'success'
        ? 'Approval recorded.'
        : `Approval failed: ${result.error ?? result.status}`,
    );
  }

  async deliver(message: {
    id: string;
    conversationId: string;
    senderId: string;
    recipientId: string;
    replyTo?: string;
    content: string;
    timestamp: string;
  }): Promise<void> {
    if (!this.state || !this.store || !this.pool) return;
    await deliverMessage(
      {
        collective: this.deps.collective,
        state: this.state,
        stateStore: this.store,
        apps: this.pool,
        options: this.options,
      },
      message,
    );
  }

  async stop(): Promise<void> {
    this.stopping = true;
    if (this.rescan) clearInterval(this.rescan);
    // Rebuild handles and logs failures before this promise settles.
    await this.inFlight;
    await this.pool?.stop();
    this.ctx = undefined;
  }
}
