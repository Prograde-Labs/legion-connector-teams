import type { Connector, ConnectorContext, ConnectorRuntimeDeps } from '@legion-collective/core';
import { resolveOptions, type ResolvedOptions, type TeamsConnectorOptions } from './options.js';

export class TeamsConnector implements Connector {
  readonly name = 'teams';
  private ctx?: ConnectorContext;
  private readonly options: ResolvedOptions;

  constructor(
    options: TeamsConnectorOptions,
    private readonly deps: ConnectorRuntimeDeps,
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
    this.ctx = ctx;
  }

  async deliver(): Promise<void> {
    // wired in Task 11
  }

  async stop(): Promise<void> {
    if (!this.ctx) return;
    this.ctx = undefined;
  }
}
