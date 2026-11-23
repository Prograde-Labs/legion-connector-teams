// src/connector.ts
import type {
  ApprovalRequestView,
  Connector,
  ConnectorContext,
  ConnectorRuntimeDeps,
  DeliveryOutcome,
} from '@legion-collective/core';
import { handleApprovalAction, sendResumed, type ResumedItem } from './approvals.js';
import { BotPool, defaultAppFactory, type CardInvokeResponse, type TeamsAppFactory } from './bot-pool.js';
import { buildDirectory, type BotEntry } from './directory.js';
import { describeError } from './errors.js';
import { findTeamsActor, readTeamsId } from './identity.js';
import { handleInbound, isBotTenant, tenantMismatchText, type TeamsContextLike } from './inbound.js';
import { deliverMessage } from './outbound.js';
import { decidedApprovalCard } from './render.js';
import { resolveOptions, type ResolvedOptions, type TeamsConnectorOptions } from './options.js';
import { SerialQueue } from './serialize.js';
import { StateStore, type TeamsState } from './state.js';

function approvalResponse(value: string): CardInvokeResponse {
  return { status: 200, body: { statusCode: 200, type: 'application/vnd.microsoft.activity.message', value } };
}

function approvalCardResponse(card: ReturnType<typeof decidedApprovalCard>): CardInvokeResponse {
  return {
    status: 200,
    body: { statusCode: 200, type: 'application/vnd.microsoft.card.adaptive', value: card },
  };
}

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
  private readonly queue = new SerialQueue();

  constructor(
    options: TeamsConnectorOptions,
    private readonly deps: ConnectorRuntimeDeps,
    private readonly appFactory: TeamsAppFactory = defaultAppFactory,
  ) {
    this.options = resolveOptions(options);
  }

  canDeliver(participantId: string): boolean {
    const participant = this.deps.collective.get(participantId);
    if (participant?.identities?.some(
      (identity) => identity.connector === 'teams' || identity.connector === 'teams-group',
    )) return true;
    return this.options.fallbackUserId === participantId;
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
    const { ctx, state, store } = this;
    if (!ctx || !state || !store) return;
    const activity = context.activity;
    const key = activity.conversation?.conversationType === 'groupChat'
      ? `${bot.appId}:${activity.conversation?.id || 'unknown'}`
      : `${bot.appId}:${readTeamsId(activity.from) || 'unknown'}`;
    await this.queue.run(key, () => handleInbound(
      {
        bot,
        ctx,
        collective: this.deps.collective,
        state,
        stateStore: store,
        options: this.options,
      },
      context,
    ));
  }

  private async onInstall(bot: BotEntry, context: TeamsContextLike): Promise<void> {
    if (!this.state || !this.store) return;
    const activity = context.activity;
    if (!isBotTenant(bot, activity)) return;
    if (activity.conversation?.conversationType !== 'personal' || !activity.conversation.id?.trim()) return;
    const aadObjectId = readTeamsId(activity.from);
    if (!aadObjectId) return;
    const participant = findTeamsActor(this.deps.collective, aadObjectId);
    if (!participant) return;
    this.state.conversations[`${bot.appId}:${participant.id}`] = {
      conversationId: activity.conversation.id,
      conversationType: 'personal',
      externalId: aadObjectId,
      tenantId: activity.channelData?.tenant?.id,
      serviceUrl: activity.serviceUrl,
      displayName: activity.from?.name,
      updatedAt: new Date().toISOString(),
    };
    await this.store.save(this.state);
  }

  private async onApproval(
    bot: BotEntry,
    context: TeamsContextLike,
    decision: 'approve' | 'reject',
  ): Promise<CardInvokeResponse> {
    if (!this.ctx) return approvalResponse('This bot is not available.');
    if (!isBotTenant(bot, context.activity)) return approvalResponse(tenantMismatchText);
    const conversationType = context.activity.conversation?.conversationType;
    if (conversationType !== 'personal' && !(conversationType === 'groupChat' && this.options.groupChats)) {
      return approvalResponse(this.options.groupChats
        ? 'Only personal chats and group chats are supported.'
        : 'Only personal chats are supported.');
    }
    const aadObjectId = readTeamsId(context.activity.from);
    const participant = aadObjectId ? findTeamsActor(this.deps.collective, aadObjectId) : undefined;
    if (!participant) return approvalResponse('Not registered.');
    const deciderName = participant.name ?? context.activity.from?.name ?? participant.id;
    const actionData = (context.activity.value?.action?.data ?? {}) as {
      approvalId?: string;
      tool?: string;
      argsText?: string;
      requesterId?: string;
      reason?: string;
    };
    const approvalId = actionData.approvalId;
    if (typeof approvalId !== 'string' || !approvalId.trim()) return approvalResponse('Missing approval id.');
    // Card input values are merged into the action data for Action.Execute.
    // Only a rejection carries a reason; ignore anything typed before approving.
    const reason =
      decision === 'reject' && typeof actionData.reason === 'string'
        ? actionData.reason.trim() || undefined
        : undefined;
    const result = await handleApprovalAction({
      ctx: this.ctx,
      userId: participant.id,
      decision,
      approvalId,
      message: reason,
    });
    if (result.status !== 'success') {
      return approvalResponse(`Approval failed: ${result.error ?? result.status}`);
    }
    const results = (result.data as {
      results?: Array<{ approvalId?: string; outcome?: string }>;
    } | undefined)?.results;
    const outcome =
      results?.find((entry) => entry.approvalId === approvalId)?.outcome ?? results?.[0]?.outcome;
    const decided =
      outcome === 'approve' || outcome === 'already_acknowledged'
        ? 'approved'
        : outcome === 'reject'
          ? 'rejected'
          : outcome === 'resume_pending' || outcome === 'resume_deferred'
            ? decision === 'approve'
              ? 'approved'
              : 'rejected'
            : undefined;
    if (!decided) {
      return approvalResponse(`Approval failed: ${outcome ?? 'unknown outcome'}`);
    }
    const resumed = (result.data as { resumed?: ResumedItem[] } | undefined)?.resumed;
    if (resumed?.length) {
      try {
        await sendResumed(resumed, (activity) => context.send(activity), this.options.aiLabel);
      } catch (error) {
        console.error('Teams resumed delivery failed', describeError(error));
      }
    }
    return approvalCardResponse(
      decidedApprovalCard({
        tool: actionData.tool,
        argsText: actionData.argsText,
        requesterId: actionData.requesterId,
        outcome: decided,
        ...(reason === undefined ? {} : { message: reason }),
        decidedBy: deciderName,
      }),
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
    approvalRequests?: ApprovalRequestView[];
  }): Promise<DeliveryOutcome> {
    if (!this.state || !this.store || !this.pool) {
      return { status: 'not-reachable', reason: 'Teams connector is not started' };
    }
    return deliverMessage(
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
