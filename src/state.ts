import { chmod, mkdir, readFile, rename, writeFile } from 'node:fs/promises';
import { join } from 'node:path';

export type ConversationType = 'personal' | 'groupChat';

export interface ConversationRecord {
  conversationId: string;
  conversationType: ConversationType;
  externalId?: string;
  name?: string;
  tenantId?: string;
  serviceUrl?: string;
  displayName?: string;
  updatedAt: string;
}

export interface PendingPair {
  appId: string;
  connector: string;
  externalId: string;
  expiresAt: string;
}

export interface MessageRecord {
  appId: string;
  conversationId: string;
  activityId: string;
  direction: 'inbound' | 'outbound';
  legionMessageId?: string;
  legionConversationId?: string;
  createdAt: string;
}

export interface TeamsState {
  version: 3;
  conversations: Record<string, ConversationRecord>;
  pendingPairs: Record<string, PendingPair>;
  legionConversations: Record<string, string>;
  legionConversationTargets: Record<string, string>;
  messages: Record<string, MessageRecord>;
  legionMessages: Record<string, string>;
}

export function emptyState(): TeamsState {
  return {
    version: 3,
    conversations: {},
    pendingPairs: {},
    legionConversations: {},
    legionConversationTargets: {},
    messages: {},
    legionMessages: {},
  };
}

let tempCounter = 0;
let saveQueue: Promise<void> = Promise.resolve();

function isPlainObject(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

interface V1ConversationRecord {
  conversationId: string;
  tenantId?: string;
  serviceUrl?: string;
  aadObjectId?: string;
  displayName?: string;
  updatedAt: string;
}

interface V1PendingPair {
  appId: string;
  aadObjectId: string;
  expiresAt: string;
}

interface V1State {
  version: 1;
  conversations: Record<string, V1ConversationRecord>;
  pendingPairs: Record<string, V1PendingPair>;
  legionConversations: Record<string, string>;
}

function isV1(value: unknown): value is V1State {
  return (
    isPlainObject(value) &&
    value.version === 1 &&
    isPlainObject(value.conversations) &&
    isPlainObject(value.pendingPairs) &&
    isPlainObject(value.legionConversations)
  );
}

interface V2State {
  version: 2;
  conversations: Record<string, ConversationRecord>;
  pendingPairs: Record<string, PendingPair>;
  legionConversations: Record<string, string>;
  legionConversationTargets?: Record<string, string>;
}

function isV2(value: unknown): value is V2State {
  return (
    isPlainObject(value) &&
    value.version === 2 &&
    isPlainObject(value.conversations) &&
    isPlainObject(value.pendingPairs) &&
    isPlainObject(value.legionConversations) &&
    (value.legionConversationTargets === undefined || isPlainObject(value.legionConversationTargets))
  );
}

function isV3(value: unknown): value is TeamsState {
  return (
    isPlainObject(value) &&
    value.version === 3 &&
    isPlainObject(value.conversations) &&
    isPlainObject(value.pendingPairs) &&
    isPlainObject(value.legionConversations) &&
    isPlainObject(value.legionConversationTargets) &&
    isPlainObject(value.messages) &&
    isPlainObject(value.legionMessages)
  );
}

function migrateV1(v1: V1State): TeamsState {
  const conversations: Record<string, ConversationRecord> = {};
  for (const [key, record] of Object.entries(v1.conversations)) {
    if (!isPlainObject(record)) continue;
    const conversation = record as unknown as V1ConversationRecord;
    conversations[key] = {
      conversationId: conversation.conversationId,
      conversationType: 'personal',
      ...(conversation.aadObjectId === undefined ? {} : { externalId: conversation.aadObjectId }),
      ...(conversation.tenantId === undefined ? {} : { tenantId: conversation.tenantId }),
      ...(conversation.serviceUrl === undefined ? {} : { serviceUrl: conversation.serviceUrl }),
      ...(conversation.displayName === undefined ? {} : { displayName: conversation.displayName }),
      updatedAt: conversation.updatedAt,
    };
  }

  const pendingPairs: Record<string, PendingPair> = {};
  for (const [code, pair] of Object.entries(v1.pendingPairs)) {
    if (!isPlainObject(pair)) continue;
    const pending = pair as unknown as V1PendingPair;
    pendingPairs[code] = {
      appId: pending.appId,
      connector: 'teams',
      externalId: pending.aadObjectId,
      expiresAt: pending.expiresAt,
    };
  }

  return {
    version: 3,
    conversations,
    pendingPairs,
    legionConversations: v1.legionConversations,
    legionConversationTargets: {},
    messages: {},
    legionMessages: {},
  };
}

function migrateV2(v2: V2State): TeamsState {
  return {
    version: 3,
    conversations: v2.conversations ?? {},
    pendingPairs: v2.pendingPairs ?? {},
    legionConversations: v2.legionConversations ?? {},
    legionConversationTargets: v2.legionConversationTargets ?? {},
    messages: {},
    legionMessages: {},
  };
}

function normalizeV3(parsed: TeamsState): TeamsState {
  return {
    version: 3,
    conversations: parsed.conversations ?? {},
    pendingPairs: parsed.pendingPairs ?? {},
    legionConversations: parsed.legionConversations ?? {},
    legionConversationTargets: parsed.legionConversationTargets ?? {},
    messages: parsed.messages ?? {},
    legionMessages: parsed.legionMessages ?? {},
  };
}

export class StateStore {
  private readonly file: string;

  constructor(private readonly dir: string) {
    this.file = join(dir, 'state.json');
  }

  async load(): Promise<TeamsState> {
    let raw: string;
    try {
      raw = await readFile(this.file, 'utf8');
    } catch (err) {
      if ((err as NodeJS.ErrnoException).code === 'ENOENT') return emptyState();
      throw err;
    }

    let parsed: unknown;
    try {
      parsed = JSON.parse(raw);
    } catch {
      await this.backup();
      return emptyState();
    }

    if (isV3(parsed)) return normalizeV3(parsed);
    if (isV2(parsed)) return migrateV2(parsed);
    if (isV1(parsed)) return migrateV1(parsed);

    await this.backup();
    return emptyState();
  }

  async save(state: TeamsState): Promise<void> {
    const run = saveQueue.then(() => this.write(state));
    saveQueue = run.catch(() => undefined);
    return run;
  }

  private async write(state: TeamsState): Promise<void> {
    await mkdir(this.dir, { recursive: true });
    const tmp = `${this.file}.${process.pid}.${tempCounter++}.tmp`;
    await writeFile(tmp, JSON.stringify(state, null, 2), { encoding: 'utf8', mode: 0o600 });
    await rename(tmp, this.file);
    await chmod(this.file, 0o600);
  }

  private async backup(): Promise<void> {
    const path = `${this.file}.bak-${Date.now()}`;
    await rename(this.file, path);
    console.warn(`[teams] Corrupt state backed up to ${path}`);
  }
}
