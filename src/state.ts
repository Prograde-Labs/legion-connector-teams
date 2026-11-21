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

export class StateStore {
  private readonly file: string;

  constructor(private readonly dir: string) {
    this.file = join(dir, 'state.json');
  }

  async load(): Promise<TeamsState> {
    try {
      const raw = await readFile(this.file, 'utf8');
      const parsed = JSON.parse(raw) as Partial<TeamsState>;
      return {
        version: 1,
        conversations: parsed.conversations ?? {},
        pendingPairs: parsed.pendingPairs ?? {},
        legionConversations: parsed.legionConversations ?? {},
      };
    } catch (err) {
      if ((err as NodeJS.ErrnoException).code === 'ENOENT') return emptyState();
      await rename(this.file, `${this.file}.bak-${Date.now()}`).catch(() => undefined);
      return emptyState();
    }
  }

  async save(state: TeamsState): Promise<void> {
    await mkdir(this.dir, { recursive: true });
    const tmp = `${this.file}.${process.pid}.tmp`;
    await writeFile(tmp, JSON.stringify(state, null, 2), { encoding: 'utf8', mode: 0o600 });
    await rename(tmp, this.file);
    await chmod(this.file, 0o600);
  }
}
