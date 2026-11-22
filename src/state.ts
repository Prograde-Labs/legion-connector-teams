import { chmod, mkdir, readFile, rename, writeFile } from 'node:fs/promises';
import { join } from 'node:path';

export interface ConversationRecord {
  conversationId: string;
  tenantId?: string;
  serviceUrl?: string;
  aadObjectId?: string;
  displayName?: string;
  updatedAt: string;
}

export interface PendingPair {
  appId: string;
  aadObjectId: string;
  expiresAt: string;
}

export interface TeamsState {
  version: 1;
  conversations: Record<string, ConversationRecord>;
  pendingPairs: Record<string, PendingPair>;
  legionConversations: Record<string, string>;
}

export function emptyState(): TeamsState {
  return { version: 1, conversations: {}, pendingPairs: {}, legionConversations: {} };
}

let tempCounter = 0;
let saveQueue: Promise<void> = Promise.resolve();

function isPlainObject(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function isState(value: unknown): value is TeamsState {
  return (
    isPlainObject(value) &&
    value.version === 1 &&
    isPlainObject(value.conversations) &&
    isPlainObject(value.pendingPairs) &&
    isPlainObject(value.legionConversations)
  );
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

    if (!isState(parsed)) {
      await this.backup();
      return emptyState();
    }

    return {
      version: 1,
      conversations: parsed.conversations,
      pendingPairs: parsed.pendingPairs,
      legionConversations: parsed.legionConversations,
    };
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
