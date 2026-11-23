# Teams Group Chats Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Treat each Microsoft Teams group chat as one Legion `user` participant (`teams-group` identity), with the same unknown-user policy, configurable sender attribution, native streaming, group-scoped commands, and approvals decided by any registered member.

**Architecture:** Add a conversation-type branch to the existing connector. Personal chats keep the current `teams` identity path; group chats resolve by `conversation.id` under a `teams-group` identity, store a per-conversation state record, prefix the sender's name before `communicate`, and route replies through a Legion-conversation target map. No Legion core changes are required.

**Tech Stack:** TypeScript ESM/NodeNext, Node 22, `@microsoft/teams.apps` 2.2.1, `@microsoft/teams.api` 2.2.1, `@microsoft/teams.cards` 2.2.1, Express 5, Vitest.

**Spec:** `docs/specs/2026-11-22-teams-group-chats-design.md`

## Global Constraints

- Node **22.x** (`node --version`). Legion's Docker image is `node:22-slim`.
- Teams SDK packages pinned to **2.2.1** (`apps`, `api`, `cards`).
- Connector name `teams`; personal identity connector `teams`; group identity connector `teams-group`; bot identity connector `teams-bot`.
- Group chats are enabled by default (`groupChats: true`); team channels remain unsupported.
- A group chat is **one** Legion participant. Do not create per-human participants.
- Existing personal-chat behavior, state keys, approvals, rejection reasons, and typing behavior must remain intact.
- State schema is **v2**. Writes are atomic, mode `0600`, and v1 files migrate on load.
- No secrets in git. Bot credentials live only in the Legion SecretStore under `teams-bot/<appId>`.
- Every task ends with its tests passing and a commit.
- Connector test command: `npm test` (Vitest). Single file: `npx vitest run src/<file>.test.ts`.
- Typecheck: `npm run typecheck`. Build: `npm run build`.
- Work in `repos/legion-connector-teams` on `main`; commit after each task.

## File Structure

**Modified**
- `src/options.ts` — `groupChats`, `groupSenderPrefix` and validation.
- `src/state.ts` — v2 types, migration, target map.
- `src/pairing.ts` — generalized pending-pair key.
- `src/identity.ts` — personal resolution plus group resolution and provisioning IDs.
- `src/commands.ts` — group-scoped commands and `/pair` as the typing human.
- `src/inbound.ts` — group branch, mention stripping, attribution, group provisioning.
- `src/connector.ts` — per-conversation serialization, group approvals, v2 install records.
- `src/approvals.ts` — unchanged behavior, group delivery comes from the caller's context.
- `src/render.ts` — decider name on decided cards.
- `src/outbound.ts` — target-map and `teams-group` routing.
- `README.md`, `docs/deploy.md` — group-chat documentation.

**Created**
- `src/attribution.ts` — sender-prefix rendering.
- `src/serialize.ts` — per-conversation promise queue.
- `src/attribution.test.ts`, `src/options.test.ts`, `src/serialize.test.ts`.

**Tests updated beside each module** (`src/*.test.ts`).

---

### Task 1: Group-chat options and attribution

**Files:**
- Modify: `src/options.ts`
- Create: `src/attribution.ts`
- Test: `src/attribution.test.ts`, `src/options.test.ts`

**Interfaces:**
- Produces:
  - `TeamsConnectorOptions.groupChats?: boolean`
  - `TeamsConnectorOptions.groupSenderPrefix?: string | null`
  - `ResolvedOptions.groupChats: boolean`
  - `ResolvedOptions.groupSenderPrefix: string | null`
  - `applyGroupPrefix(template: string | null, senderName: string, message: string): string`

- [ ] **Step 1: Write the failing attribution tests**

Create `src/attribution.test.ts`:

```ts
import { describe, expect, it } from 'vitest';
import { applyGroupPrefix } from './attribution.js';

describe('applyGroupPrefix', () => {
  it('prefixes the sender name by default', () => {
    expect(applyGroupPrefix('[{name}]: {message}', 'Alice', 'hello')).toBe('[Alice]: hello');
  });

  it('supports a custom template', () => {
    expect(applyGroupPrefix('{name} says: {message}', 'Alice', 'hello')).toBe('Alice says: hello');
  });

  it('leaves the message unchanged when attribution is disabled', () => {
    expect(applyGroupPrefix(null, 'Alice', 'hello')).toBe('hello');
  });

  it('works when the template has no name placeholder', () => {
    expect(applyGroupPrefix('{message}', 'Alice', 'hello')).toBe('hello');
  });
});
```

- [ ] **Step 2: Run the attribution tests to verify they fail**

Run: `npx vitest run src/attribution.test.ts`
Expected: FAIL because `src/attribution.ts` does not exist.

- [ ] **Step 3: Implement attribution**

Create `src/attribution.ts`:

```ts
export function applyGroupPrefix(
  template: string | null,
  senderName: string,
  message: string,
): string {
  if (template === null) return message;
  return template.replaceAll('{name}', senderName).replace('{message}', message);
}
```

- [ ] **Step 4: Write the failing options tests**

Create `src/options.test.ts`:

```ts
import { describe, expect, it } from 'vitest';
import { resolveOptions } from './options.js';

describe('resolveOptions group chat options', () => {
  it('defaults group chats on with the default attribution template', () => {
    const options = resolveOptions({ stateDir: '/tmp/x' });
    expect(options.groupChats).toBe(true);
    expect(options.groupSenderPrefix).toBe('[{name}]: {message}');
  });

  it('accepts a null attribution template', () => {
    expect(resolveOptions({ stateDir: '/tmp/x', groupSenderPrefix: null }).groupSenderPrefix).toBeNull();
  });

  it('accepts a custom attribution template', () => {
    expect(resolveOptions({ stateDir: '/tmp/x', groupSenderPrefix: '{message} — {name}' }).groupSenderPrefix)
      .toBe('{message} — {name}');
  });

  it('rejects a template without a message placeholder', () => {
    expect(() => resolveOptions({ stateDir: '/tmp/x', groupSenderPrefix: '{name}' }))
      .toThrow(/\{message\}/);
  });

  it('can disable group chats', () => {
    expect(resolveOptions({ stateDir: '/tmp/x', groupChats: false }).groupChats).toBe(false);
  });
});
```

- [ ] **Step 5: Run the options tests to verify they fail**

Run: `npx vitest run src/options.test.ts`
Expected: FAIL because `groupChats` and `groupSenderPrefix` are absent.

- [ ] **Step 6: Add the options**

Modify `src/options.ts`. Add to `TeamsConnectorOptions`:

```ts
  groupChats?: boolean;
  groupSenderPrefix?: string | null;
```

Add to `ResolvedOptions`:

```ts
  groupChats: boolean;
  groupSenderPrefix: string | null;
```

In `resolveOptions`, before `return`:

```ts
  const groupSenderPrefix = options.groupSenderPrefix === undefined
    ? '[{name}]: {message}'
    : options.groupSenderPrefix;
  if (groupSenderPrefix !== null && !groupSenderPrefix.includes('{message}')) {
    throw new Error('groupSenderPrefix must contain {message}');
  }
```

Add to the returned object:

```ts
    groupChats: options.groupChats ?? true,
    groupSenderPrefix,
```

- [ ] **Step 7: Run the tests and typecheck**

Run: `npx vitest run src/attribution.test.ts src/options.test.ts`
Expected: PASS.

Run: `npm run typecheck`
Expected: PASS.

- [ ] **Step 8: Commit**

```bash
git add src/options.ts src/attribution.ts src/attribution.test.ts src/options.test.ts
git commit -m "feat(options): add group chat options and attribution"
```

---

### Task 2: State schema v2 and generalized pending pairs

**Files:**
- Modify: `src/state.ts`
- Modify: `src/pairing.ts`
- Modify: `src/identity.ts`
- Modify: `src/commands.ts`
- Modify: `src/inbound.ts`
- Modify: `src/connector.ts`
- Modify: `src/outbound.ts`
- Test: `src/state.test.ts`, `src/pairing.test.ts`, `src/commands.test.ts`, `src/inbound.test.ts`, `src/outbound.test.ts`

**Interfaces:**
- Consumes: nothing from earlier tasks.
- Produces:
  - `ConversationType = 'personal' | 'groupChat'`
  - `ConversationRecord` with `conversationType`, `externalId`, `name`, `displayName`, `tenantId`, `serviceUrl`, `updatedAt`
  - `PendingPair` with `appId`, `connector`, `externalId`, `expiresAt`
  - `TeamsState` with `version: 2`, `conversations`, `pendingPairs`, `legionConversations`, `legionConversationTargets`
  - `emptyState(): TeamsState`
  - `PendingPairKey = { appId: string; connector: string; externalId: string }`
  - `rememberPendingPair(state: TeamsState, key: PendingPairKey, ttlSeconds: number, now?: number): string`
  - `validatePendingPair(state: TeamsState, code: string, now?: number): PendingPair | null`
  - `consumePendingPair(state: TeamsState, code: string, now?: number): PendingPair | null`

- [ ] **Step 1: Write the failing state tests**

Replace `src/state.test.ts` with:

```ts
import { chmod, mkdir, mkdtemp, readdir, readFile, stat, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { describe, expect, it, vi } from 'vitest';
import { emptyState, StateStore } from './state.js';

const isRoot = typeof process.getuid === 'function' && process.getuid() === 0;

async function backups(dir: string): Promise<string[]> {
  return (await readdir(dir)).filter((name) => name.startsWith('state.json.bak-'));
}

describe('StateStore', () => {
  it('returns empty v2 state when the file is missing', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'teams-state-'));
    await expect(new StateStore(dir).load()).resolves.toEqual(emptyState());
  });

  it('round-trips v2 state', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'teams-state-'));
    const store = new StateStore(dir);
    const state = emptyState();
    state.conversations['app:alice'] = {
      conversationId: 'conv-1',
      conversationType: 'personal',
      externalId: 'aad-1',
      updatedAt: '2026-11-21T00:00:00.000Z',
    };
    state.conversations['app:group:19:g@thread.v2'] = {
      conversationId: '19:g@thread.v2',
      conversationType: 'groupChat',
      externalId: '19:g@thread.v2',
      name: 'Engineering',
      updatedAt: '2026-11-21T00:00:00.000Z',
    };
    state.legionConversationTargets['thread-1'] = 'app:group:19:g@thread.v2';
    await store.save(state);
    await expect(store.load()).resolves.toEqual(state);
  });

  it('writes the file with mode 0600', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'teams-state-'));
    const store = new StateStore(dir);
    await store.save(emptyState());
    const mode = (await stat(join(dir, 'state.json'))).mode & 0o777;
    expect(mode).toBe(0o600);
  });

  it('serializes overlapping saves', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'teams-state-'));
    const store = new StateStore(dir);
    const first = emptyState();
    first.legionConversations['a'] = 'legion-a';
    const second = emptyState();
    second.legionConversations['b'] = 'legion-b';

    await expect(Promise.all([store.save(first), store.save(second)])).resolves.toHaveLength(2);

    await expect(store.load()).resolves.toEqual(second);
  });

  it('migrates v1 conversations and pending pairs', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'teams-state-'));
    await writeFile(join(dir, 'state.json'), JSON.stringify({
      version: 1,
      conversations: {
        'app:alice': {
          conversationId: 'conv-1',
          tenantId: 'tenant-1',
          serviceUrl: 'https://smba.trafficmanager.net/teams/',
          aadObjectId: 'aad-1',
          displayName: 'Alice',
          updatedAt: '2026-11-21T00:00:00.000Z',
        },
      },
      pendingPairs: {
        ABC234: {
          appId: 'app',
          aadObjectId: 'aad-2',
          expiresAt: '2026-11-22T00:00:00.000Z',
        },
      },
      legionConversations: { 'app:alice': 'thread-1' },
    }));
    await expect(new StateStore(dir).load()).resolves.toEqual({
      version: 2,
      conversations: {
        'app:alice': {
          conversationId: 'conv-1',
          conversationType: 'personal',
          externalId: 'aad-1',
          tenantId: 'tenant-1',
          serviceUrl: 'https://smba.trafficmanager.net/teams/',
          displayName: 'Alice',
          updatedAt: '2026-11-21T00:00:00.000Z',
        },
      },
      pendingPairs: {
        ABC234: {
          appId: 'app',
          connector: 'teams',
          externalId: 'aad-2',
          expiresAt: '2026-11-22T00:00:00.000Z',
        },
      },
      legionConversations: { 'app:alice': 'thread-1' },
      legionConversationTargets: {},
    });
  });

  it('defaults a missing target map on load', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'teams-state-'));
    await writeFile(join(dir, 'state.json'), JSON.stringify({
      version: 2,
      conversations: {},
      pendingPairs: {},
      legionConversations: {},
    }));
    await expect(new StateStore(dir).load()).resolves.toEqual(emptyState());
  });

  it('backs up corrupt state and starts fresh', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'teams-state-'));
    await writeFile(join(dir, 'state.json'), '{not json');
    const store = new StateStore(dir);
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    try {
      await expect(store.load()).resolves.toEqual(emptyState());
      const [saved] = await backups(dir);
      expect(warn).toHaveBeenCalledWith(`[teams] Corrupt state backed up to ${join(dir, saved)}`);
      expect(JSON.stringify(warn.mock.calls)).not.toContain('{not json');
    } finally { warn.mockRestore(); }
    const raw = await readFile(join(dir, 'state.json'), 'utf8').catch(() => null);
    expect(raw).toBeNull();
    const saved = await backups(dir);
    expect(saved).toHaveLength(1);
    await expect(readFile(join(dir, saved[0]), 'utf8')).resolves.toBe('{not json');
  });

  it.each([
    [
      'non-object map',
      '{"version":2,"conversations":[],"pendingPairs":{},"legionConversations":{},"legionConversationTargets":{}}',
    ],
    [
      'unsupported version',
      '{"version":3,"conversations":{},"pendingPairs":{},"legionConversations":{}}',
    ],
  ])('backs up valid JSON with malformed shape: %s', async (_label, contents) => {
    const dir = await mkdtemp(join(tmpdir(), 'teams-state-'));
    await writeFile(join(dir, 'state.json'), contents);
    const store = new StateStore(dir);
    await expect(store.load()).resolves.toEqual(emptyState());
    const raw = await readFile(join(dir, 'state.json'), 'utf8').catch(() => null);
    expect(raw).toBeNull();
    const saved = await backups(dir);
    expect(saved).toHaveLength(1);
    await expect(readFile(join(dir, saved[0]), 'utf8')).resolves.toBe(contents);
  });

  it('propagates read errors other than ENOENT', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'teams-state-'));
    await mkdir(join(dir, 'state.json'));
    await expect(new StateStore(dir).load()).rejects.toMatchObject({ code: 'EISDIR' });
  });

  it.skipIf(isRoot)('propagates backup failures', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'teams-state-'));
    await writeFile(join(dir, 'state.json'), '{not json');
    await chmod(dir, 0o500);
    try {
      await expect(new StateStore(dir).load()).rejects.toMatchObject({ code: 'EACCES' });
    } finally {
      await chmod(dir, 0o700);
    }
  });
});
```

- [ ] **Step 2: Run the state tests to verify they fail**

Run: `npx vitest run src/state.test.ts`
Expected: FAIL because `emptyState()` is still v1 and there is no migration or target map.

- [ ] **Step 3: Implement state v2**

Replace `src/state.ts` with:

```ts
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

export interface TeamsState {
  version: 2;
  conversations: Record<string, ConversationRecord>;
  pendingPairs: Record<string, PendingPair>;
  legionConversations: Record<string, string>;
  legionConversationTargets: Record<string, string>;
}

export function emptyState(): TeamsState {
  return {
    version: 2,
    conversations: {},
    pendingPairs: {},
    legionConversations: {},
    legionConversationTargets: {},
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

function isV2(value: unknown): value is TeamsState {
  return (
    isPlainObject(value) &&
    value.version === 2 &&
    isPlainObject(value.conversations) &&
    isPlainObject(value.pendingPairs) &&
    isPlainObject(value.legionConversations) &&
    (value.legionConversationTargets === undefined || isPlainObject(value.legionConversationTargets))
  );
}

function migrateV1(v1: V1State): TeamsState {
  const conversations: Record<string, ConversationRecord> = {};
  for (const [key, record] of Object.entries(v1.conversations)) {
    conversations[key] = {
      conversationId: record.conversationId,
      conversationType: 'personal',
      ...(record.aadObjectId === undefined ? {} : { externalId: record.aadObjectId }),
      ...(record.tenantId === undefined ? {} : { tenantId: record.tenantId }),
      ...(record.serviceUrl === undefined ? {} : { serviceUrl: record.serviceUrl }),
      ...(record.displayName === undefined ? {} : { displayName: record.displayName }),
      updatedAt: record.updatedAt,
    };
  }

  const pendingPairs: Record<string, PendingPair> = {};
  for (const [code, pair] of Object.entries(v1.pendingPairs)) {
    pendingPairs[code] = {
      appId: pair.appId,
      connector: 'teams',
      externalId: pair.aadObjectId,
      expiresAt: pair.expiresAt,
    };
  }

  return {
    version: 2,
    conversations,
    pendingPairs,
    legionConversations: v1.legionConversations,
    legionConversationTargets: {},
  };
}

function normalizeV2(parsed: Record<string, unknown>): TeamsState {
  return {
    version: 2,
    conversations: parsed.conversations as Record<string, ConversationRecord>,
    pendingPairs: parsed.pendingPairs as Record<string, PendingPair>,
    legionConversations: parsed.legionConversations as Record<string, string>,
    legionConversationTargets: (parsed.legionConversationTargets as Record<string, string>) ?? {},
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

    if (isV2(parsed)) return normalizeV2(parsed);
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
```

- [ ] **Step 4: Run the state tests**

Run: `npx vitest run src/state.test.ts`
Expected: PASS.

- [ ] **Step 5: Update the pairing tests**

Replace every `rememberPendingPair(state, 'app', 'aad', 600, 1000)` call in `src/pairing.test.ts` with:

```ts
rememberPendingPair(state, { appId: 'app', connector: 'teams', externalId: 'aad' }, 600, 1000)
```

Add this test:

```ts
  it('keeps codes separate by connector and external id', () => {
    const state = emptyState();
    const personal = rememberPendingPair(
      state, { appId: 'app', connector: 'teams', externalId: 'aad-1' }, 600, 1000,
    );
    const group = rememberPendingPair(
      state, { appId: 'app', connector: 'teams-group', externalId: 'aad-1' }, 600, 1000,
    );
    expect(group).not.toBe(personal);
    expect(state.pendingPairs[personal]).toMatchObject({ connector: 'teams', externalId: 'aad-1' });
    expect(state.pendingPairs[group]).toMatchObject({ connector: 'teams-group', externalId: 'aad-1' });
  });
```

Update the malformed-entry fixtures:

```ts
state.pendingPairs['CODE'] = { appId: 'app', connector: 'teams', externalId: 'aad', expiresAt: 123 as never };
state.pendingPairs['ABC234'] = { appId: 'app', connector: 'teams', externalId: 'aad', expiresAt };
state.pendingPairs['INVALID'] = { appId: 'app', connector: 'teams', externalId: 'aad', expiresAt: 'not-a-date' };
```

- [ ] **Step 6: Run the pairing tests to verify they fail**

Run: `npx vitest run src/pairing.test.ts`
Expected: FAIL because `rememberPendingPair` still takes an `aadObjectId` string.

- [ ] **Step 7: Implement the generalized pending pair**

Replace `src/pairing.ts` with:

```ts
import { randomInt } from 'node:crypto';
import type { PendingPair, TeamsState } from './state.js';

const ALPHABET = '23456789ABCDEFGHJKLMNPQRSTUVWXYZ';

export interface PendingPairKey {
  appId: string;
  connector: string;
  externalId: string;
}

export function generatePairCode(random: () => number = () => randomInt(0, ALPHABET.length)): string {
  let code = '';
  for (let i = 0; i < 6; i++) code += ALPHABET[random()];
  return code;
}

export function pruneExpiredPairs(state: TeamsState, now = Date.now()): void {
  for (const [code, pair] of Object.entries(state.pendingPairs)) {
    const expiry = pair && typeof pair === 'object' && typeof pair.expiresAt === 'string'
      ? Date.parse(pair.expiresAt) : NaN;
    if (Number.isNaN(expiry) || expiry <= now) delete state.pendingPairs[code];
  }
}

export function rememberPendingPair(
  state: TeamsState,
  key: PendingPairKey,
  ttlSeconds: number,
  now = Date.now(),
): string {
  pruneExpiredPairs(state, now);
  for (const [code, pair] of Object.entries(state.pendingPairs)) {
    if (pair.appId === key.appId && pair.connector === key.connector && pair.externalId === key.externalId) {
      return code;
    }
  }
  let code = generatePairCode();
  while (state.pendingPairs[code]) code = generatePairCode();
  state.pendingPairs[code] = {
    appId: key.appId,
    connector: key.connector,
    externalId: key.externalId,
    expiresAt: new Date(now + ttlSeconds * 1000).toISOString(),
  };
  return code;
}

export function validatePendingPair(
  state: TeamsState,
  code: string,
  now = Date.now(),
): PendingPair | null {
  const key = code.toUpperCase();
  const pair = state.pendingPairs[key];
  if (!Object.hasOwn(state.pendingPairs, key)) return null;
  const expiry = pair && typeof pair === 'object' && typeof pair.expiresAt === 'string'
    ? Date.parse(pair.expiresAt) : NaN;
  if (Number.isNaN(expiry) || expiry <= now) {
    delete state.pendingPairs[key];
    return null;
  }
  return pair;
}

export function consumePendingPair(
  state: TeamsState,
  code: string,
  now = Date.now(),
): PendingPair | null {
  const pair = validatePendingPair(state, code, now);
  if (pair) delete state.pendingPairs[code.toUpperCase()];
  return pair;
}
```

- [ ] **Step 8: Update personal identity resolution**

In `src/identity.ts`, replace the pairing case in `resolveIdentity` with:

```ts
    case 'pairing':
      return {
        kind: 'pairing',
        aadObjectId,
        code: rememberPendingPair(
          state,
          { appId, connector: 'teams', externalId: aadObjectId },
          options.pairingTtlSeconds,
          now,
        ),
      };
```

- [ ] **Step 9: Update commands to use the generalized pending pair**

In `src/commands.ts`, replace the `pair` call and success message:

```ts
        const { result } = await deps.ctx.callTool(caller, 'pair', {
          participantId: target,
          connector: pending.connector,
          externalId: pending.externalId,
        });
        if (result.status === 'success') {
          delete deps.state.pendingPairs[code.toUpperCase()];
          const label = pending.connector === 'teams'
            ? 'Teams'
            : pending.connector === 'teams-group'
              ? 'Teams group'
              : pending.connector;
          await deps.send(`Paired ${target} with ${label} identity ${pending.externalId}.`);
        } else if (isAwaitingApproval(result)) {
```

- [ ] **Step 10: Update the command test fixtures**

In `src/commands.test.ts`, replace every `rememberPendingPair(d.state, 'app', 'aad-1', 600)` with:

```ts
rememberPendingPair(d.state, { appId: 'app', connector: 'teams', externalId: 'aad-1' }, 600)
```

Replace `rememberPendingPair(d.state, 'app', 'aad-9', 600)` with:

```ts
rememberPendingPair(d.state, { appId: 'app', connector: 'teams', externalId: 'aad-9' }, 600)
```

Replace `rememberPendingPair(d.state, 'other-app', 'aad-1', 600)` with:

```ts
rememberPendingPair(d.state, { appId: 'other-app', connector: 'teams', externalId: 'aad-1' }, 600)
```

Update the malformed-entry fixture:

```ts
d.state.pendingPairs['ABC234'] = {
  appId: 'app', connector: 'teams', externalId: 'aad-1', expiresAt,
};
```

- [ ] **Step 11: Update the inbound fallback fixture**

In `src/inbound.test.ts`, replace:

```ts
const code = rememberPendingPair(s.state, 'app', 'aad-7', 600);
```

with:

```ts
const code = rememberPendingPair(
  s.state, { appId: 'app', connector: 'teams', externalId: 'aad-7' }, 600,
);
```

- [ ] **Step 12: Update every `ConversationRecord` writer to v2 fields**

In `src/inbound.ts`, change `upsertConversation` to:

```ts
  state.conversations[key] = {
    conversationId: activity.conversation?.id ?? '',
    conversationType: 'personal',
    externalId: aadObjectId,
    tenantId: activity.channelData?.tenant?.id,
    serviceUrl: activity.serviceUrl,
    displayName,
    updatedAt: new Date().toISOString(),
  };
```

In `src/connector.ts`, change the `onInstall` record to:

```ts
    this.state.conversations[`${bot.appId}:${participant.id}`] = {
      conversationId: activity.conversation.id,
      conversationType: 'personal',
      externalId: aadObjectId,
      tenantId: activity.channelData?.tenant?.id,
      serviceUrl: activity.serviceUrl,
      displayName: activity.from?.name,
      updatedAt: new Date().toISOString(),
    };
```

In `src/outbound.ts`, change the personal conversation record to:

```ts
      record = {
        conversationId: conversation.id,
        conversationType: 'personal',
        externalId: userIdentity.externalId,
        tenantId: deps.options.tenantId,
        updatedAt: new Date().toISOString(),
      };
```

In `src/outbound.test.ts`, update the personal creation assertion:

```ts
    expect(s.state.conversations['app:alice']).toEqual(expect.objectContaining({
      conversationId: 'conv-new',
      conversationType: 'personal',
      externalId: 'aad-1',
      tenantId: 'tenant-1',
    }));
```

- [ ] **Step 13: Run the affected tests**

Run: `npx vitest run src/state.test.ts src/pairing.test.ts src/commands.test.ts src/identity.test.ts src/inbound.test.ts src/outbound.test.ts`
Expected: PASS.

- [ ] **Step 14: Run the full suite and typecheck**

Run: `npm test`
Expected: PASS.

Run: `npm run typecheck`
Expected: PASS.

- [ ] **Step 15: Commit**

```bash
git add src/state.ts src/pairing.ts src/identity.ts src/commands.ts src/inbound.ts src/connector.ts src/outbound.ts \
  src/state.test.ts src/pairing.test.ts src/commands.test.ts src/inbound.test.ts src/outbound.test.ts
git commit -m "feat(state): add v2 schema and generalized pending pairs"
```

---

### Task 3: Group identity resolution and provisioning IDs

**Files:**
- Modify: `src/identity.ts`
- Test: `src/identity.test.ts`

**Interfaces:**
- Consumes: `PendingPairKey`, `rememberPendingPair`, and `TeamsState` from Task 2.
- Produces:
  - `ResolvedGroup = { participantId: string; conversationId: string; name?: string }`
  - `GroupIdentityOutcome` (resolved / pairing / reject / auto-provision / fallback)
  - `findGroupActor(collective, conversationId)`
  - `resolveGroupIdentity(args): GroupIdentityOutcome`
  - `provisionedGroupId(name: string | undefined, conversationId: string): string`
  - `provisionedUserId(name, aadObjectId)` (moved from `inbound.ts`)
  - `hash6(value: string): string`

- [ ] **Step 1: Write the failing group identity tests**

Append to `src/identity.test.ts`:

```ts
import { findGroupActor, provisionedGroupId, provisionedUserId, resolveGroupIdentity } from './identity.js';

describe('group identity', () => {
  const groupIdentity = (id: string, conversationId: string) => ({
    id,
    name: 'Engineering',
    type: 'user',
    tools: {},
    status: 'active',
    identities: [{ connector: 'teams-group', externalId: conversationId }],
  });

  it('resolves a registered group by conversation id', () => {
    const outcome = resolveGroupIdentity({
      collective: collective([groupIdentity('team-eng', '19:g@thread.v2')]),
      state: emptyState(),
      appId: 'app',
      conversationId: '19:g@thread.v2',
      name: 'Engineering',
      options: options(),
    });
    expect(outcome).toEqual({
      kind: 'resolved',
      group: { participantId: 'team-eng', conversationId: '19:g@thread.v2', name: 'Engineering' },
    });
  });

  it('issues a group pairing code for an unknown group', () => {
    const state = emptyState();
    const outcome = resolveGroupIdentity({
      collective: collective([]),
      state,
      appId: 'app',
      conversationId: '19:g@thread.v2',
      options: options(),
    });
    expect(outcome.kind).toBe('pairing');
    expect(Object.values(state.pendingPairs)[0]).toMatchObject({
      appId: 'app',
      connector: 'teams-group',
      externalId: '19:g@thread.v2',
    });
  });

  it('returns group auto-provision for that policy', () => {
    const outcome = resolveGroupIdentity({
      collective: collective([]),
      state: emptyState(),
      appId: 'app',
      conversationId: '19:g@thread.v2',
      name: 'Engineering',
      options: options({ unknownUserPolicy: 'auto-provision' }),
    });
    expect(outcome).toEqual({
      kind: 'auto-provision',
      externalId: '19:g@thread.v2',
      name: 'Engineering',
    });
  });

  it('rejects an unknown group for the reject policy', () => {
    const outcome = resolveGroupIdentity({
      collective: collective([]),
      state: emptyState(),
      appId: 'app',
      conversationId: '19:g@thread.v2',
      options: options({ unknownUserPolicy: 'reject' }),
    });
    expect(outcome).toEqual({ kind: 'reject', externalId: '19:g@thread.v2' });
  });

  it('uses the fallback participant for an unknown group', () => {
    const outcome = resolveGroupIdentity({
      collective: collective([
        { id: 'operator', name: 'Operator', type: 'user', tools: {}, status: 'active' },
      ]),
      state: emptyState(),
      appId: 'app',
      conversationId: '19:g@thread.v2',
      name: 'Engineering',
      options: options({ unknownUserPolicy: 'fallback', fallbackUserId: 'operator' }),
    });
    expect(outcome).toEqual({
      kind: 'fallback',
      group: { participantId: 'operator', conversationId: '19:g@thread.v2', name: 'Engineering' },
    });
  });

  it('treats a retired group participant as unknown', () => {
    const outcome = resolveGroupIdentity({
      collective: collective([
        { ...groupIdentity('team-old', '19:g@thread.v2'), status: 'retired' },
      ]),
      state: emptyState(),
      appId: 'app',
      conversationId: '19:g@thread.v2',
      options: options({ unknownUserPolicy: 'reject' }),
    });
    expect(outcome.kind).toBe('reject');
  });

  it('builds stable, readable provisioning ids', () => {
    expect(provisionedGroupId('Engineering', '19:group@thread.v2'))
      .toBe('teams-group-engineering-eqrgnh');
    expect(provisionedGroupId(undefined, '19:group@thread.v2'))
      .toBe('teams-group-group-eqrgnh');
    expect(provisionedUserId('Alice', 'aad-1')).toBe('teams-alice-1j42js');
  });

  it('finds a group actor by conversation id', () => {
    const found = findGroupActor(collective([groupIdentity('team-eng', '19:g@thread.v2')]), '19:g@thread.v2');
    expect(found?.id).toBe('team-eng');
  });
});
```

- [ ] **Step 2: Run the identity tests to verify they fail**

Run: `npx vitest run src/identity.test.ts`
Expected: FAIL because the group functions do not exist.

- [ ] **Step 3: Implement group identity**

Modify `src/identity.ts`. Add after the existing `findTeamsActor`:

```ts
export interface ResolvedGroup {
  participantId: string;
  conversationId: string;
  name?: string;
}

export type GroupIdentityOutcome =
  | { kind: 'resolved'; group: ResolvedGroup }
  | { kind: 'pairing'; externalId: string; code: string }
  | { kind: 'reject'; externalId: string }
  | { kind: 'auto-provision'; externalId: string; name?: string }
  | { kind: 'fallback'; group: ResolvedGroup };

function slug(value: string): string {
  return value.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '').slice(0, 24) || 'user';
}

export function hash6(value: string): string {
  let hash = 0;
  for (const char of value) hash = (hash * 31 + char.charCodeAt(0)) >>> 0;
  return hash.toString(36).padStart(6, '0').slice(0, 6);
}

export function provisionedUserId(name: string, aadObjectId: string): string {
  return `teams-${slug(name)}-${hash6(aadObjectId)}`;
}

export function provisionedGroupId(name: string | undefined, conversationId: string): string {
  const base = name?.trim() ? name : 'group';
  return `teams-group-${slug(base)}-${hash6(conversationId)}`;
}

export function findGroupActor(collective: CollectiveLike, conversationId: string) {
  const participant = collective.findByIdentity('teams-group', conversationId);
  return participant?.status === 'retired' ? undefined : participant;
}

export function resolveGroupIdentity(args: {
  collective: CollectiveLike;
  state: TeamsState;
  appId: string;
  conversationId: string;
  name?: string;
  options: ResolvedOptions;
  now?: number;
}): GroupIdentityOutcome {
  const { collective, state, appId, conversationId, name, options, now } = args;

  const found = findGroupActor(collective, conversationId);
  if (found) {
    return { kind: 'resolved', group: { participantId: found.id, conversationId, name } };
  }

  if (options.unknownUserPolicy === 'fallback' && options.fallbackUserId) {
    const fallback = collective.get(options.fallbackUserId);
    if (fallback && (fallback.status ?? 'active') !== 'retired') {
      return { kind: 'fallback', group: { participantId: fallback.id, conversationId, name } };
    }
    return { kind: 'reject', externalId: conversationId };
  }

  switch (options.unknownUserPolicy) {
    case 'pairing':
      return {
        kind: 'pairing',
        externalId: conversationId,
        code: rememberPendingPair(
          state,
          { appId, connector: 'teams-group', externalId: conversationId },
          options.pairingTtlSeconds,
          now,
        ),
      };
    case 'auto-provision':
      return { kind: 'auto-provision', externalId: conversationId, name };
    case 'reject':
    case 'fallback':
      return { kind: 'reject', externalId: conversationId };
  }
}
```

- [ ] **Step 4: Remove the duplicate helpers from `src/inbound.ts`**

Delete `slug`, `hash6`, and `provisionedUserId` from `src/inbound.ts`, and import `provisionedUserId` from `./identity.js` where needed. (The full inbound rewrite in Task 5 will also use `provisionedGroupId` and `hash6`.)

- [ ] **Step 5: Run the identity tests**

Run: `npx vitest run src/identity.test.ts`
Expected: PASS.

- [ ] **Step 6: Commit**

```bash
git add src/identity.ts src/identity.test.ts src/inbound.ts
git commit -m "feat(identity): resolve Teams group chats by conversation id"
```

---

### Task 4: Group-scoped commands and `/pair` as the typing human

**Files:**
- Modify: `src/commands.ts`
- Test: `src/commands.test.ts`

**Interfaces:**
- Consumes: `PendingPair.connector` / `externalId` and `legionConversationTargets` from Task 2.
- Produces:
  - `CommandDeps` adds `senderParticipantId?`, `senderName?`, `conversationId?`, `isGroup?`
  - `/new` clears the group key and its target-map entry.
  - `/whoami` shows group participant, group conversation ID, sender, and bot app.
  - `/pair` calls `pair` as `pairingParticipantId ?? senderParticipantId`, using the pending `connector` and `externalId`.

- [ ] **Step 1: Update the existing command tests for the new pending-pair shape**

In `src/commands.test.ts`, replace every `rememberPendingPair(d.state, 'app', 'aad-1', 600)` with:

```ts
rememberPendingPair(d.state, { appId: 'app', connector: 'teams', externalId: 'aad-1' }, 600)
```

Replace `rememberPendingPair(d.state, 'app', 'aad-9', 600)` with:

```ts
rememberPendingPair(d.state, { appId: 'app', connector: 'teams', externalId: 'aad-9' }, 600)
```

Replace `rememberPendingPair(d.state, 'other-app', 'aad-1', 600)` with:

```ts
rememberPendingPair(d.state, { appId: 'other-app', connector: 'teams', externalId: 'aad-1' }, 600)
```

Update the malformed-entry fixture:

```ts
d.state.pendingPairs['ABC234'] = {
  appId: 'app', connector: 'teams', externalId: 'aad-1', expiresAt,
};
```

- [ ] **Step 2: Add failing group command tests**

Append to `src/commands.test.ts`:

```ts
  it('clears the group conversation and its target map entry on /new', async () => {
    const { deps: d, sent } = deps('/new', {
      participantId: 'team-eng',
      isGroup: true,
      conversationKey: 'app:team-eng:19:g@thread.v2',
      conversationId: '19:g@thread.v2',
    });
    d.state.legionConversations['app:team-eng:19:g@thread.v2'] = 'legion-1';
    d.state.legionConversationTargets['legion-1'] = 'app:team-eng:19:g@thread.v2';
    await expect(handleCommand(d)).resolves.toBe(true);
    expect(d.state.legionConversations['app:team-eng:19:g@thread.v2']).toBeUndefined();
    expect(d.state.legionConversationTargets['legion-1']).toBeUndefined();
    expect(sent[0]).toMatch(/new conversation/i);
  });

  it('shows group identity details on /whoami', async () => {
    const { deps: d, sent } = deps('/whoami', {
      participantId: 'team-eng',
      senderName: 'Alice',
      isGroup: true,
      conversationKey: 'app:team-eng:19:g@thread.v2',
      conversationId: '19:g@thread.v2',
    });
    await expect(handleCommand(d)).resolves.toBe(true);
    expect(sent[0]).toContain('Participant: team-eng');
    expect(sent[0]).toContain('Group conversation ID: 19:g@thread.v2');
    expect(sent[0]).toContain('Sender: Alice');
    expect(sent[0]).toContain('Bot app: app');
  });

  it('runs /pair as the typing human in a group and uses the pending group identity', async () => {
    const { deps: d, callTool, sent } = deps('/pair', {
      participantId: 'team-eng',
      senderParticipantId: 'alice',
      isGroup: true,
      conversationKey: 'app:team-eng:19:g@thread.v2',
      conversationId: '19:g@thread.v2',
    });
    const code = rememberPendingPair(
      d.state,
      { appId: 'app', connector: 'teams-group', externalId: '19:g@thread.v2' },
      600,
    );
    await handleCommand({ ...d, text: `/pair ${code} team-eng` });
    expect(callTool).toHaveBeenCalledWith('alice', 'pair', {
      participantId: 'team-eng',
      connector: 'teams-group',
      externalId: '19:g@thread.v2',
    });
    expect(sent[0]).toMatch(/Paired team-eng with Teams group identity 19:g@thread\.v2/);
    expect(d.state.pendingPairs[code]).toBeUndefined();
  });

  it('rejects /pair from an unregistered human in a group', async () => {
    const { deps: d, callTool, sent } = deps('/pair', {
      participantId: 'team-eng',
      isGroup: true,
      conversationKey: 'app:team-eng:19:g@thread.v2',
      conversationId: '19:g@thread.v2',
    });
    const code = rememberPendingPair(
      d.state,
      { appId: 'app', connector: 'teams-group', externalId: '19:g@thread.v2' },
      600,
    );
    await handleCommand({ ...d, text: `/pair ${code} team-eng` });
    expect(callTool).not.toHaveBeenCalled();
    expect(sent[0]).toBe('Only registered users can redeem pairing codes.');
    expect(d.state.pendingPairs[code]).toBeDefined();
  });
```

- [ ] **Step 3: Run the command tests to verify they fail**

Run: `npx vitest run src/commands.test.ts`
Expected: FAIL because `senderParticipantId`, `isGroup`, `conversationId`, and `senderName` are not handled.

- [ ] **Step 4: Implement the command changes**

Replace `src/commands.ts` with:

```ts
// src/commands.ts
import type { ConnectorContext } from '@legion-collective/core';
import { isAwaitingApproval } from './approvals.js';
import { validatePendingPair } from './pairing.js';
import type { ResolvedOptions } from './options.js';
import type { TeamsState } from './state.js';

const pairingInFlight = new Set<string>();

export interface CommandDeps {
  appId: string;
  participantId: string;
  senderParticipantId?: string;
  aadObjectId: string;
  senderName?: string;
  conversationKey?: string;
  conversationId?: string;
  isFallback?: boolean;
  isGroup?: boolean;
  text: string;
  state: TeamsState;
  ctx: ConnectorContext;
  options: ResolvedOptions;
  send: (text: string) => Promise<void>;
}

export async function handleCommand(deps: CommandDeps): Promise<boolean> {
  const [raw, ...rest] = deps.text.trim().split(/\s+/);
  const command = raw.toLowerCase();
  const key = deps.conversationKey ?? `${deps.appId}:${deps.participantId}`;

  switch (command) {
    case '/new': {
      const previous = deps.state.legionConversations[key];
      delete deps.state.legionConversations[key];
      if (previous) delete deps.state.legionConversationTargets[previous];
      await deps.send('Started a new conversation.');
      return true;
    }
    case '/pair': {
      if (deps.isFallback) {
        await deps.send('Only registered users can redeem pairing codes.');
        return true;
      }
      const callerId = deps.senderParticipantId ?? (deps.isGroup ? undefined : deps.participantId);
      if (!callerId) {
        await deps.send('Only registered users can redeem pairing codes.');
        return true;
      }
      const code = rest[0]?.toUpperCase();
      const target = rest[1];
      if (!code || !target) {
        await deps.send('Usage: /pair <code> <participant-id>');
        return true;
      }
      if (pairingInFlight.has(code)) {
        await deps.send('Pairing already in progress.');
        return true;
      }
      pairingInFlight.add(code);
      try {
        const pending = validatePendingPair(deps.state, code);
        if (!pending || pending.appId !== deps.appId) {
          await deps.send('Pairing code is invalid or expired.');
          return true;
        }
        const caller = deps.options.pairingParticipantId ?? callerId;
        const { result } = await deps.ctx.callTool(caller, 'pair', {
          participantId: target,
          connector: pending.connector,
          externalId: pending.externalId,
        });
        if (result.status === 'success') {
          delete deps.state.pendingPairs[code.toUpperCase()];
          const label = pending.connector === 'teams'
            ? 'Teams'
            : pending.connector === 'teams-group'
              ? 'Teams group'
              : pending.connector;
          await deps.send(`Paired ${target} with ${label} identity ${pending.externalId}.`);
        } else if (isAwaitingApproval(result)) {
          await deps.send('Pairing is awaiting approval.');
        } else {
          await deps.send(`Pairing failed: ${result.error ?? result.status}`);
        }
        return true;
      } finally {
        pairingInFlight.delete(code);
      }
    }
    case '/whoami': {
      const conversation = deps.state.legionConversations[key] ?? '(none)';
      if (deps.isGroup) {
        await deps.send(
          `Participant: ${deps.participantId}\n` +
          `Group conversation ID: ${deps.conversationId ?? '(unknown)'}\n` +
          `Sender: ${deps.senderName ?? deps.aadObjectId}\n` +
          `Bot app: ${deps.appId}\nConversation: ${conversation}`,
        );
      } else {
        await deps.send(
          `Participant: ${deps.participantId}\nTeams ID: ${deps.aadObjectId}\n` +
          `Bot app: ${deps.appId}\nConversation: ${conversation}`,
        );
      }
      return true;
    }
    case '/help': {
      await deps.send(
        'Commands:\n' +
        '/new — start a new conversation\n' +
        '/pair <code> <participant-id> — pair a pending Teams identity\n' +
        '/whoami — show your identity\n' +
        '/help — this message',
      );
      return true;
    }
    default:
      return false;
  }
}
```

- [ ] **Step 5: Run the command tests**

Run: `npx vitest run src/commands.test.ts`
Expected: PASS.

- [ ] **Step 6: Commit**

```bash
git add src/commands.ts src/commands.test.ts
git commit -m "feat(commands): group-scoped commands and human-run pairing"
```

---

### Task 5: Group inbound path

**Files:**
- Modify: `src/inbound.ts`
- Test: `src/inbound.test.ts`

**Interfaces:**
- Consumes:
  - `resolveGroupIdentity`, `findGroupActor`, `findTeamsActor`, `provisionedGroupId`, `hash6`, `provisionedUserId` from Task 3.
  - `applyGroupPrefix` and `ResolvedOptions.groupChats` / `groupSenderPrefix` from Task 1.
  - `ConversationRecord`, `TeamsState` from Task 2.
  - `handleCommand` with `senderParticipantId`, `isGroup`, `conversationId`, `senderName` from Task 4.
- Produces:
  - `handleInbound` supports `conversationType === 'groupChat'` when `options.groupChats` is true.
  - Group state keys `<appId>:<participantId>:<conversationId>`.
  - `legionConversationTargets[legionConversationId] = platformKey` after every communicate turn.
  - Group provisioning with `connector: 'teams-group'`.
  - Bare group mentions send an empty message.

- [ ] **Step 1: Update the existing inbound tests for v2 and group support**

In `src/inbound.test.ts`, update the unsupported-chat cases:

```ts
  it.each([
    [{ conversation: { conversationType: 'channel', id: 'channel-1' } }, 'Only personal chats and group chats are supported.'],
    [{ conversation: { id: 'conv-1' } }, 'Only personal chats and group chats are supported.'],
    [{ channelData: { tenant: { id: 'other' } } }, 'This bot is not available in your tenant.'],
    [{ channelData: undefined }, 'This bot is not available in your tenant.'],
  ] as const)('rejects unsupported chat or tenant data before resolving identity', async (override, text) => {
    const s = setup([alice]);
    const send = captureReply();
    await handleInbound(inboundDeps(s), { activity: { ...activity('hello'), ...override }, send });
    expect(send).toHaveBeenCalledOnce();
    expect(send).toHaveBeenCalledWith(expect.objectContaining({ text }));
    expect(s.collective.findByIdentity).not.toHaveBeenCalled();
    expect(s.save).not.toHaveBeenCalled();
  });
```

Add a disabled-group case:

```ts
  it('rejects group chats when groupChats is disabled', async () => {
    const s = setup([alice]);
    const send = captureReply();
    await handleInbound(
      inboundDeps(s, { groupChats: false }),
      { activity: groupActivity('hello'), send },
    );
    expect(send).toHaveBeenCalledWith(expect.objectContaining({ text: 'Only personal chats are supported.' }));
    expect(s.collective.findByIdentity).not.toHaveBeenCalled();
  });
```

Confirm the fallback test in `src/inbound.test.ts` uses the object form from Task 2:

```ts
const code = rememberPendingPair(
  s.state, { appId: 'app', connector: 'teams', externalId: 'aad-7' }, 600,
);
```

Add the group activity helper near the personal `activity` helper:

```ts
const groupActivity = (text: string, overrides: Record<string, unknown> = {}) => ({
  from: { id: '29:1', name: 'Alice', aadObjectId: 'aad-1' },
  recipient: { id: 'bot-1' },
  conversation: { id: '19:group@thread.v2', conversationType: 'groupChat', name: 'Engineering' },
  channelData: { tenant: { id: 't' } },
  serviceUrl: 'https://smba.trafficmanager.net/teams/',
  entities: [
    { type: 'mention', text: '<at>Assistant</at>', mentioned: { id: 'bot-1', name: 'Assistant' } },
  ],
  text,
  ...overrides,
});
```

- [ ] **Step 2: Add failing group inbound tests**

Append to `src/inbound.test.ts`:

```ts
describe('handleInbound group chats', () => {
  const group = {
    id: 'team-eng', name: 'Engineering', type: 'user', tools: {}, status: 'active',
    identities: [{ connector: 'teams-group', externalId: '19:group@thread.v2' }],
  };

  it('communicates as the group participant with attribution', async () => {
    const s = setup([alice, group]);
    const send = captureReply();
    await handleInbound(inboundDeps(s), { activity: groupActivity('hello'), send });
    const key = 'app:team-eng:19:group@thread.v2';
    expect(s.streamTool).toHaveBeenCalledWith('team-eng', 'communicate', expect.objectContaining({
      to: 'assistant',
      message: '[Alice]: hello',
      conversationId: undefined,
    }));
    expect(s.state.conversations[key]).toMatchObject({
      conversationId: '19:group@thread.v2',
      conversationType: 'groupChat',
      externalId: '19:group@thread.v2',
      name: 'Engineering',
    });
    expect(s.state.legionConversations[key]).toBe('thread-1');
    expect(s.state.legionConversationTargets['thread-1']).toBe(key);
    expect(s.ctx.registry.setActive).toHaveBeenCalledWith('team-eng', 'teams');
  });

  it('removes only the bot mention and preserves human mentions', async () => {
    const s = setup([alice, group]);
    const send = captureReply();
    const activity = groupActivity('<at>Assistant</at> <at>Bob</at> please review', {
      entities: [
        { type: 'mention', text: '<at>Assistant</at>', mentioned: { id: 'bot-1', name: 'Assistant' } },
        { type: 'mention', text: '<at>Bob</at>', mentioned: { id: 'human-2', name: 'Bob' } },
      ],
    });
    await handleInbound(inboundDeps(s), { activity, send });
    expect(s.streamTool).toHaveBeenCalledWith('team-eng', 'communicate', expect.objectContaining({
      message: '[Alice]: <at>Bob</at> please review',
    }));
  });

  it('sends an empty message for a bare mention', async () => {
    const s = setup([alice, group]);
    const send = captureReply();
    await handleInbound(inboundDeps(s), { activity: groupActivity('<at>Assistant</at>'), send });
    expect(s.streamTool).toHaveBeenCalledWith('team-eng', 'communicate', expect.objectContaining({
      message: '',
    }));
    expect(send).not.toHaveBeenCalledWith(expect.objectContaining({ text: 'Empty message.' }));
  });

  it('can disable attribution', async () => {
    const s = setup([alice, group]);
    await handleInbound(
      inboundDeps(s, { groupSenderPrefix: null }),
      { activity: groupActivity('hello'), send: captureReply() },
    );
    expect(s.streamTool).toHaveBeenCalledWith('team-eng', 'communicate', expect.objectContaining({
      message: 'hello',
    }));
  });

  it('issues a pairing code for an unknown group', async () => {
    const s = setup([alice]);
    const send = captureReply();
    await handleInbound(inboundDeps(s), { activity: groupActivity('hello'), send });
    const [code, pending] = Object.entries(s.state.pendingPairs)[0];
    expect(pending).toMatchObject({
      appId: 'app', connector: 'teams-group', externalId: '19:group@thread.v2',
    });
    expect(send).toHaveBeenCalledWith(expect.objectContaining({
      text: expect.stringMatching(/isn't registered.*Group ID: 19:group@thread\.v2/),
    }));
    expect(s.streamTool).not.toHaveBeenCalled();
    expect(code).toMatch(/^[2-9A-HJ-NP-Z]{6}$/);
  });

  it('tells a rejected group how to register', async () => {
    const s = setup([alice]);
    const send = captureReply();
    await handleInbound(
      inboundDeps(s, { unknownUserPolicy: 'reject' }),
      { activity: groupActivity('hello'), send },
    );
    expect(send).toHaveBeenCalledWith(expect.objectContaining({
      text: expect.stringMatching(/19:group@thread\.v2.*administrator/i),
    }));
  });

  it('auto-provisions and pairs an unknown group', async () => {
    const s = setup([alice]);
    const send = captureReply();
    await handleInbound(
      inboundDeps(s, { unknownUserPolicy: 'auto-provision' }),
      { activity: groupActivity('hello'), send },
    );
    expect(s.callTool).toHaveBeenCalledWith('operator', 'create_user', expect.objectContaining({
      id: 'teams-group-engineering-eqrgnh',
      name: 'Engineering',
      tools: { communicate: 'auto', approval_response: 'auto' },
    }));
    expect(s.callTool).toHaveBeenCalledWith('operator', 'pair', {
      participantId: 'teams-group-engineering-eqrgnh',
      connector: 'teams-group',
      externalId: '19:group@thread.v2',
    });
    expect(send.mock.calls[0][0]).toHaveProperty('text', expect.stringMatching(/Welcome/));
    expect(s.streamTool).not.toHaveBeenCalled();
  });

  it('isolates fallback group conversations by conversation id', async () => {
    const s = setup([alice, { id: 'operator', name: 'Operator', type: 'user', tools: {}, status: 'active' }]);
    const deps = inboundDeps(s, { unknownUserPolicy: 'fallback', fallbackUserId: 'operator' });
    const send = captureReply();
    s.streamTool
      .mockResolvedValueOnce({
        conversationId: 'conv-1',
        gen: (async function* (): AsyncGenerator<StreamChunk> {
          yield { type: 'text_delta', delta: 'Hi' };
          yield { type: 'stream:done', result: { status: 'success', data: { conversationId: 'thread-g1' } } };
        })(),
      })
      .mockResolvedValueOnce({
        conversationId: 'conv-1',
        gen: (async function* (): AsyncGenerator<StreamChunk> {
          yield { type: 'text_delta', delta: 'Hi' };
          yield { type: 'stream:done', result: { status: 'success', data: { conversationId: 'thread-g2' } } };
        })(),
      });
    for (const conversationId of ['19:g1@thread.v2', '19:g2@thread.v2']) {
      await handleInbound(deps, {
        activity: groupActivity('hello', {
          conversation: { id: conversationId, conversationType: 'groupChat', name: conversationId },
        }),
        send,
      });
    }
    expect(s.streamTool).toHaveBeenNthCalledWith(1, 'operator', 'communicate', expect.objectContaining({
      message: '[Alice]: hello',
    }));
    expect(s.state.conversations).toMatchObject({
      'app:operator:19:g1@thread.v2': { conversationId: '19:g1@thread.v2' },
      'app:operator:19:g2@thread.v2': { conversationId: '19:g2@thread.v2' },
    });
    expect(s.state.legionConversationTargets).toEqual({
      'thread-g1': 'app:operator:19:g1@thread.v2',
      'thread-g2': 'app:operator:19:g2@thread.v2',
    });
  });

  it('handles /new on the group key without calling communicate', async () => {
    const s = setup([alice, group]);
    const key = 'app:team-eng:19:group@thread.v2';
    s.state.legionConversations[key] = 'old';
    s.state.legionConversationTargets['old'] = key;
    const send = captureReply();
    await handleInbound(inboundDeps(s), { activity: groupActivity('/new'), send });
    expect(s.state.legionConversations[key]).toBeUndefined();
    expect(s.state.legionConversationTargets['old']).toBeUndefined();
    expect(s.streamTool).not.toHaveBeenCalled();
  });

  it('runs /pair as the typing human in a group', async () => {
    const s = setup([alice, group]);
    const code = rememberPendingPair(
      s.state,
      { appId: 'app', connector: 'teams-group', externalId: '19:other@thread.v2' },
      600,
    );
    await handleInbound(inboundDeps(s), {
      activity: groupActivity(`/pair ${code} team-other`),
      send: captureReply(),
    });
    expect(s.callTool).toHaveBeenCalledWith('alice', 'pair', {
      participantId: 'team-other',
      connector: 'teams-group',
      externalId: '19:other@thread.v2',
    });
    expect(s.state.pendingPairs[code]).toBeUndefined();
  });

  it('rejects /pair from an unregistered human in a group', async () => {
    const s = setup([group]);
    const code = rememberPendingPair(
      s.state,
      { appId: 'app', connector: 'teams-group', externalId: '19:other@thread.v2' },
      600,
    );
    const send = captureReply();
    await handleInbound(inboundDeps(s), {
      activity: groupActivity(`/pair ${code} team-other`),
      send,
    });
    expect(s.callTool).not.toHaveBeenCalled();
    expect(send).toHaveBeenCalledWith(expect.objectContaining({
      text: 'Only registered users can redeem pairing codes.',
    }));
    expect(s.state.pendingPairs[code]).toBeDefined();
  });
});
```

- [ ] **Step 3: Run the inbound tests to verify they fail**

Run: `npx vitest run src/inbound.test.ts`
Expected: FAIL because group chats are rejected and the helpers are absent.

- [ ] **Step 4: Implement the group inbound path**

Replace `src/inbound.ts` with:

```ts
// src/inbound.ts
import type { ConnectorContext } from '@legion-collective/core';
import type { CollectiveLike } from './directory.js';
import { MessageActivityInput, TypingActivityInput, stripMentionsText, type IMessageActivity } from '@microsoft/teams.api';
import { applyGroupPrefix } from './attribution.js';
import { handleCommand } from './commands.js';
import { isAwaitingApproval } from './approvals.js';
import { runCommunicate, type TeamsStreamer } from './communicate.js';
import type { BotEntry } from './directory.js';
import {
  findTeamsActor,
  hash6,
  provisionedGroupId,
  provisionedUserId,
  readTeamsId,
  resolveGroupIdentity,
  resolveIdentity,
  type ResolvedGroup,
} from './identity.js';
import type { ResolvedOptions } from './options.js';
import { sendText } from './render.js';
import type { StateStore, TeamsState } from './state.js';

export interface TeamsActivity {
  from?: { id?: string; name?: string; aadObjectId?: string };
  recipient?: { id?: string };
  conversation?: { id?: string; name?: string; conversationType?: string };
  channelData?: { tenant?: { id?: string } };
  serviceUrl?: string;
  text?: string;
  entities?: Array<{
    type?: string;
    text?: string;
    mentioned?: { id?: string; name?: string };
  }>;
  value?: { action?: { data?: Record<string, unknown> } };
}

export interface TeamsContextLike {
  activity: TeamsActivity;
  send: (activity: MessageActivityInput | TypingActivityInput | string) => Promise<unknown>;
  stream?: TeamsStreamer;
}

export function isBotTenant(bot: BotEntry, activity: TeamsActivity): boolean {
  return activity.channelData?.tenant?.id === bot.credentials.tenantId;
}

export const tenantMismatchText = 'This bot is not available in your tenant.';

export interface InboundDeps {
  bot: BotEntry;
  ctx: ConnectorContext;
  collective: CollectiveLike;
  state: TeamsState;
  stateStore: StateStore;
  options: ResolvedOptions;
}

function upsertPersonalConversation(
  state: TeamsState,
  key: string,
  externalId: string,
  displayName: string | undefined,
  activity: TeamsActivity,
): void {
  state.conversations[key] = {
    conversationId: activity.conversation?.id ?? '',
    conversationType: 'personal',
    externalId,
    tenantId: activity.channelData?.tenant?.id,
    serviceUrl: activity.serviceUrl,
    displayName,
    updatedAt: new Date().toISOString(),
  };
}

function upsertGroupConversation(
  state: TeamsState,
  key: string,
  group: ResolvedGroup,
  activity: TeamsActivity,
): void {
  state.conversations[key] = {
    conversationId: group.conversationId,
    conversationType: 'groupChat',
    externalId: group.conversationId,
    name: group.name,
    tenantId: activity.channelData?.tenant?.id,
    serviceUrl: activity.serviceUrl,
    updatedAt: new Date().toISOString(),
  };
}

function stripBotMention(activity: TeamsActivity): string {
  return stripMentionsText(activity as unknown as IMessageActivity, {
    accountId: activity.recipient?.id,
  })?.trim() ?? '';
}

async function autoProvision(
  deps: InboundDeps,
  aadObjectId: string,
  displayName: string | undefined,
  reply: (text: string) => Promise<void>,
): Promise<void> {
  const name = displayName ?? 'Teams user';
  const id = provisionedUserId(name, aadObjectId);
  const existing = deps.collective.get(id);

  if (existing && (existing.type !== 'user' || existing.status === 'retired' || existing.identities?.some(
    (identity) => identity.connector === 'teams' && identity.externalId !== aadObjectId,
  ))) {
    await reply(`Participant ${id} is already in use; ask an administrator to pair you.`);
    return;
  }

  let provisioned;
  const tools = { ...existing?.tools, communicate: 'auto', approval_response: 'auto' };
  try {
    provisioned = existing
      ? await deps.ctx.callTool(deps.options.provisionerId, 'modify_user', { id, tools })
      : await deps.ctx.callTool(deps.options.provisionerId, 'create_user', { id, name, tools });
  } catch (error) {
    await reply(`Provisioning failed: ${error instanceof Error ? error.message : String(error)}`);
    return;
  }
  if (isAwaitingApproval(provisioned.result)) {
    await reply('Provisioning approval is pending.');
    return;
  }
  if (provisioned.result.status !== 'success') {
    await reply(`Provisioning failed: ${provisioned.result.error ?? provisioned.result.status}`);
    return;
  }

  let paired;
  try {
    paired = await deps.ctx.callTool(
      deps.options.pairingParticipantId ?? deps.options.provisionerId,
      'pair',
      { participantId: id, connector: 'teams', externalId: aadObjectId },
    );
  } catch (error) {
    await reply(`Pairing failed: ${error instanceof Error ? error.message : String(error)}`);
    return;
  }
  if (isAwaitingApproval(paired.result)) {
    await reply('Pairing approval is pending.');
    return;
  }
  if (paired.result.status !== 'success') {
    await reply(`Pairing failed: ${paired.result.error ?? paired.result.status}`);
    return;
  }

  await reply(`Welcome, ${name}. You are now paired as ${id}.`);
}

async function autoProvisionGroup(
  deps: InboundDeps,
  conversationId: string,
  conversationName: string | undefined,
  reply: (text: string) => Promise<void>,
): Promise<void> {
  const id = provisionedGroupId(conversationName, conversationId);
  const name = conversationName?.trim() || `Teams group ${hash6(conversationId)}`;
  const existing = deps.collective.get(id);

  if (existing && (existing.type !== 'user' || existing.status === 'retired' || existing.identities?.some(
    (identity) => identity.connector === 'teams-group' && identity.externalId !== conversationId,
  ))) {
    await reply(`Participant ${id} is already in use; ask an administrator to pair this group.`);
    return;
  }

  let provisioned;
  const tools = { ...existing?.tools, communicate: 'auto', approval_response: 'auto' };
  try {
    provisioned = existing
      ? await deps.ctx.callTool(deps.options.provisionerId, 'modify_user', { id, tools })
      : await deps.ctx.callTool(deps.options.provisionerId, 'create_user', { id, name, tools });
  } catch (error) {
    await reply(`Provisioning failed: ${error instanceof Error ? error.message : String(error)}`);
    return;
  }
  if (isAwaitingApproval(provisioned.result)) {
    await reply('Provisioning approval is pending.');
    return;
  }
  if (provisioned.result.status !== 'success') {
    await reply(`Provisioning failed: ${provisioned.result.error ?? provisioned.result.status}`);
    return;
  }

  let paired;
  try {
    paired = await deps.ctx.callTool(
      deps.options.pairingParticipantId ?? deps.options.provisionerId,
      'pair',
      { participantId: id, connector: 'teams-group', externalId: conversationId },
    );
  } catch (error) {
    await reply(`Pairing failed: ${error instanceof Error ? error.message : String(error)}`);
    return;
  }
  if (isAwaitingApproval(paired.result)) {
    await reply('Pairing approval is pending.');
    return;
  }
  if (paired.result.status !== 'success') {
    await reply(`Pairing failed: ${paired.result.error ?? paired.result.status}`);
    return;
  }

  await reply(`Welcome, ${name}. This group is now paired as ${id}.`);
}

async function runTurn(args: {
  deps: InboundDeps;
  context: TeamsContextLike;
  participantId: string;
  key: string;
  text: string;
}): Promise<void> {
  const { deps, context, participantId, key, text } = args;
  let typing: NodeJS.Timeout | undefined;
  const stopTyping = () => {
    if (typing) {
      clearInterval(typing);
      typing = undefined;
    }
  };
  if (deps.options.typingIntervalMs > 0) {
    const tick = () => context.send(new TypingActivityInput()).catch(() => undefined);
    tick();
    typing = setInterval(tick, deps.options.typingIntervalMs);
  }

  try {
    const { conversationId } = await runCommunicate({
      ctx: deps.ctx,
      userId: participantId,
      botOwnerId: deps.bot.participantId,
      text,
      conversationId: deps.state.legionConversations[key],
      options: deps.options,
      stream: context.stream,
      send: (activity) => context.send(activity),
      stopTyping,
    });
    if (conversationId !== undefined) {
      deps.state.legionConversations[key] = conversationId;
      deps.state.legionConversationTargets[conversationId] = key;
      await deps.stateStore.save(deps.state);
    }
  } finally {
    stopTyping();
  }
}

export async function handleInbound(
  deps: InboundDeps,
  context: TeamsContextLike,
): Promise<void> {
  const activity = context.activity;
  const reply = (text: string) => sendText((message) => context.send(message), deps.options.aiLabel, text);
  const conversationType = activity.conversation?.conversationType;
  const isGroup = conversationType === 'groupChat';

  if (conversationType !== 'personal' && !(isGroup && deps.options.groupChats)) {
    await reply(deps.options.groupChats
      ? 'Only personal chats and group chats are supported.'
      : 'Only personal chats are supported.');
    return;
  }
  if (!isBotTenant(deps.bot, activity)) {
    await reply(tenantMismatchText);
    return;
  }
  const aadObjectId = readTeamsId(activity.from);
  if (!aadObjectId) {
    await reply('Unable to read your Teams identity.');
    return;
  }

  if (isGroup) {
    const conversationId = activity.conversation?.id?.trim();
    if (!conversationId) {
      await reply('Unable to read the group conversation.');
      return;
    }

    const outcome = resolveGroupIdentity({
      collective: deps.collective,
      state: deps.state,
      appId: deps.bot.appId,
      conversationId,
      name: activity.conversation?.name,
      options: deps.options,
    });

    switch (outcome.kind) {
      case 'pairing':
        await deps.stateStore.save(deps.state);
        await reply(
          `This group isn't registered with Legion yet. Ask someone with the \`pair\` tool to run ` +
          `/pair ${outcome.code} <participant-id>. Group ID: ${conversationId}.`,
        );
        return;

      case 'reject':
        await reply(
          `This group isn't registered with Legion yet. Group ID: ${conversationId}. ` +
          'Ask an administrator to register it.',
        );
        return;

      case 'auto-provision':
        await autoProvisionGroup(deps, conversationId, outcome.name, reply);
        await deps.stateStore.save(deps.state);
        return;

      case 'resolved':
      case 'fallback': {
        const group = outcome.group;
        const key = `${deps.bot.appId}:${group.participantId}:${conversationId}`;
        upsertGroupConversation(deps.state, key, group, activity);
        await deps.stateStore.save(deps.state);
        deps.ctx.registry.setActive(group.participantId, 'teams');

        const text = stripBotMention(activity);
        if (text.startsWith('/')) {
          const sender = findTeamsActor(deps.collective, aadObjectId);
          const handled = await handleCommand({
            appId: deps.bot.appId,
            participantId: group.participantId,
            senderParticipantId: sender?.id,
            aadObjectId,
            senderName: activity.from?.name,
            conversationKey: key,
            conversationId,
            isFallback: outcome.kind === 'fallback',
            isGroup: true,
            text,
            state: deps.state,
            ctx: deps.ctx,
            options: deps.options,
            send: reply,
          });
          if (handled) {
            await deps.stateStore.save(deps.state);
            return;
          }
        }

        const message = text
          ? applyGroupPrefix(deps.options.groupSenderPrefix, activity.from?.name ?? 'Unknown', text)
          : '';
        await runTurn({
          deps,
          context,
          participantId: group.participantId,
          key,
          text: message,
        });
        return;
      }
    }
  }

  const outcome = resolveIdentity({
    collective: deps.collective,
    state: deps.state,
    appId: deps.bot.appId,
    aadObjectId,
    displayName: activity.from?.name,
    options: deps.options,
  });

  switch (outcome.kind) {
    case 'pairing':
      await deps.stateStore.save(deps.state);
      await reply(
        `Not registered with Legion yet. Ask someone with the \`pair\` tool to run ` +
        `/pair ${outcome.code} <participant-id>. Your Teams ID: ${outcome.aadObjectId}.`,
      );
      return;

    case 'reject':
      await reply(
        `Not registered with Legion yet. Your Teams ID: ${outcome.aadObjectId}. Ask an administrator to register you.`,
      );
      return;

    case 'auto-provision':
      await autoProvision(deps, outcome.aadObjectId, outcome.displayName, reply);
      await deps.stateStore.save(deps.state);
      return;

    case 'resolved':
    case 'fallback': {
      const user = outcome.user;
      const key = `${deps.bot.appId}:${user.participantId}` +
        (outcome.kind === 'fallback' ? `:${user.aadObjectId}` : '');
      upsertPersonalConversation(deps.state, key, user.aadObjectId, user.displayName, activity);
      await deps.stateStore.save(deps.state);
      deps.ctx.registry.setActive(user.participantId, 'teams');

      const text = activity.text?.trim();
      if (!text) {
        await reply('Empty message.');
        return;
      }

      if (text.startsWith('/')) {
        const handled = await handleCommand({
          appId: deps.bot.appId,
          participantId: user.participantId,
          senderParticipantId: user.participantId,
          aadObjectId: user.aadObjectId,
          senderName: user.displayName,
          conversationKey: key,
          isFallback: outcome.kind === 'fallback',
          text,
          state: deps.state,
          ctx: deps.ctx,
          options: deps.options,
          send: reply,
        });
        if (handled) {
          await deps.stateStore.save(deps.state);
          return;
        }
      }

      await runTurn({
        deps,
        context,
        participantId: user.participantId,
        key,
        text,
      });
      return;
    }
  }
}
```

- [ ] **Step 5: Run the inbound tests**

Run: `npx vitest run src/inbound.test.ts`
Expected: PASS.

- [ ] **Step 6: Run the full suite**

Run: `npm test`
Expected: PASS. The connector approval test still expects group chats to be rejected until Task 7 updates it.

- [ ] **Step 7: Commit**

```bash
git add src/inbound.ts src/inbound.test.ts
git commit -m "feat(inbound): treat Teams group chats as participants"
```

---

### Task 6: Per-conversation serialization

**Files:**
- Create: `src/serialize.ts`
- Create: `src/serialize.test.ts`
- Modify: `src/connector.ts`

**Interfaces:**
- Consumes: nothing from earlier tasks.
- Produces:
  - `SerialQueue.run(key: string, task: () => Promise<void>): Promise<void>`
  - Connector queues `handleInbound` by platform conversation key.

- [ ] **Step 1: Write the failing queue tests**

Create `src/serialize.test.ts`:

```ts
import { describe, expect, it } from 'vitest';
import { SerialQueue } from './serialize.js';

describe('SerialQueue', () => {
  it('serializes tasks for the same key', async () => {
    const queue = new SerialQueue();
    const order: string[] = [];
    let release!: () => void;
    const gate = new Promise<void>((resolve) => { release = resolve; });
    const first = queue.run('chat-1', async () => {
      order.push('first-start');
      await gate;
      order.push('first-end');
    });
    const second = queue.run('chat-1', async () => { order.push('second'); });

    await Promise.resolve();
    expect(order).toEqual(['first-start']);
    release();
    await Promise.all([first, second]);
    expect(order).toEqual(['first-start', 'first-end', 'second']);
  });

  it('runs tasks for different keys concurrently', async () => {
    const queue = new SerialQueue();
    const order: string[] = [];
    let release!: () => void;
    const gate = new Promise<void>((resolve) => { release = resolve; });
    const first = queue.run('chat-1', async () => {
      order.push('first-start');
      await gate;
      order.push('first-end');
    });
    const second = queue.run('chat-2', async () => { order.push('second'); });

    await Promise.resolve();
    expect(order).toEqual(['first-start', 'second']);
    release();
    await Promise.all([first, second]);
    expect(order).toEqual(['first-start', 'second', 'first-end']);
  });

  it('continues after a rejected task', async () => {
    const queue = new SerialQueue();
    await expect(queue.run('chat-1', async () => { throw new Error('boom'); })).rejects.toThrow('boom');
    await expect(queue.run('chat-1', async () => undefined)).resolves.toBeUndefined();
  });
});
```

- [ ] **Step 2: Run the queue tests to verify they fail**

Run: `npx vitest run src/serialize.test.ts`
Expected: FAIL because `src/serialize.ts` does not exist.

- [ ] **Step 3: Implement the queue**

Create `src/serialize.ts`:

```ts
export class SerialQueue {
  private readonly tails = new Map<string, Promise<void>>();

  run(key: string, task: () => Promise<void>): Promise<void> {
    const previous = this.tails.get(key) ?? Promise.resolve();
    const run = previous.then(task);
    const tail = run.then(() => undefined, () => undefined);
    this.tails.set(key, tail);
    void tail.finally(() => {
      if (this.tails.get(key) === tail) this.tails.delete(key);
    });
    return run;
  }
}
```

- [ ] **Step 4: Wire the queue into the connector**

Modify `src/connector.ts`:

```ts
import { SerialQueue } from './serialize.js';
```

Add the field:

```ts
  private readonly queue = new SerialQueue();
```

Replace `onMessage` with:

```ts
  private async onMessage(bot: BotEntry, context: TeamsContextLike): Promise<void> {
    if (!this.ctx || !this.state || !this.store) return;
    const activity = context.activity;
    const key = activity.conversation?.conversationType === 'groupChat'
      ? `${bot.appId}:${activity.conversation?.id ?? 'unknown'}`
      : `${bot.appId}:${readTeamsId(activity.from) ?? 'unknown'}`;
    await this.queue.run(key, () => handleInbound(
      {
        bot,
        ctx: this.ctx!,
        collective: this.deps.collective,
        state: this.state!,
        stateStore: this.store!,
        options: this.options,
      },
      context,
    ));
  }
```

The non-null assertions are needed because the fields are optional and TypeScript cannot narrow them across the closure. Alternatively capture them in locals:

```ts
    const { ctx, state, store } = this;
    if (!ctx || !state || !store) return;
    ...
    await this.queue.run(key, () => handleInbound(
      { bot, ctx, collective: this.deps.collective, state, stateStore: store, options: this.options },
      context,
    ));
```

Use the local-capture version to avoid `!`.

- [ ] **Step 5: Run the queue tests and connector tests**

Run: `npx vitest run src/serialize.test.ts src/connector.test.ts`
Expected: PASS. The connector approval test still expects group chats to be rejected until Task 7 updates it.

- [ ] **Step 6: Commit**

```bash
git add src/serialize.ts src/serialize.test.ts src/connector.ts
git commit -m "feat(connector): serialize inbound turns per conversation"
```

---

### Task 7: Group approvals and decider name

**Files:**
- Modify: `src/render.ts`
- Modify: `src/connector.ts`
- Test: `src/render.test.ts`, `src/connector.test.ts`

**Interfaces:**
- Consumes: `SerialQueue` and connector changes from Task 6; `ResolvedOptions.groupChats` from Task 1; `findTeamsActor` from Task 3.
- Produces:
  - `DecidedApprovalView.decidedBy?: string`
  - `decidedApprovalCard` adds a `Decided by` fact.
  - `onApproval` allows `personal` and `groupChat` (when enabled), rejects channels with `Only personal chats and group chats are supported.`
  - Install records use v2 fields.

- [ ] **Step 1: Write the failing render test**

In `src/render.test.ts`, update the rejected-card test to include a decider:

```ts
  it('builds a rejected approval card that shows the reason and decider', () => {
    const card = decidedApprovalCard({
      outcome: 'rejected',
      tool: 'file_write',
      argsText: '{"path":"test.txt"}',
      requesterId: 'assistant',
      message: 'Not on a Friday',
      decidedBy: 'Alice',
    });
    expect(card.actions ?? []).toHaveLength(0);
    expect(JSON.stringify(card)).toContain('Rejected');
    expect(JSON.stringify(card)).toContain('Not on a Friday');
    expect(JSON.stringify(card)).toContain('Decided by');
    expect(JSON.stringify(card)).toContain('Alice');
  });
```

- [ ] **Step 2: Run the render test to verify it fails**

Run: `npx vitest run src/render.test.ts`
Expected: FAIL because `decidedBy` is not rendered.

- [ ] **Step 3: Implement the decider fact**

In `src/render.ts`, add to `DecidedApprovalView`:

```ts
  decidedBy?: string;
```

In `decidedApprovalCard`, after the requester fact:

```ts
  if (view.decidedBy) facts.push({ title: 'Decided by', value: view.decidedBy });
```

- [ ] **Step 4: Update the connector approval tests**

In `src/connector.test.ts`, update the install test expectation for the v2 record:

```ts
      const personal = await record();
      expect(personal).toMatchObject({
        conversationId: 'personal-id',
        conversationType: 'personal',
        externalId: 'aad-1',
      });
```

In the approval test, replace the group rejection with a channel rejection and add a group approval case. Use this shape:

```ts
      const invoke = (id: string, conversationType = 'personal') => handlers.onApproval(entry, { send, activity: {
        from: { aadObjectId: 'aad-1', name: 'Alice' }, channelData: { tenant: { id } },
        conversation: { conversationType },
        value: { action: { data: {
          approvalId: 'ap-1',
          tool: 'file_write',
          argsText: '{"path":"test.txt"}',
          requesterId: 'assistant',
        } } },
      } } as never, 'approve');
      expect(await invoke('other')).toHaveProperty('body.value', 'This bot is not available in your tenant.');
      expect(await invoke('t', 'channel')).toHaveProperty(
        'body.value', 'Only personal chats and group chats are supported.',
      );
      expect(callTool).not.toHaveBeenCalled();
      const approved = await invoke('t', 'groupChat');
      expect(approved).toMatchObject({ status: 200, body: {
        statusCode: 200, type: 'application/vnd.microsoft.card.adaptive',
      } });
      expect(JSON.stringify(approved.body)).toContain('Approved');
      expect(JSON.stringify(approved.body)).toContain('Decided by');
      expect(JSON.stringify(approved.body)).toContain('Alice');
      expect((approved.body as { value?: { actions?: unknown[] } }).value?.actions ?? []).toHaveLength(0);
      participants[1].status = 'retired';
      expect(await invoke('t')).toHaveProperty('body.value', 'Not registered.');
      expect(callTool).toHaveBeenCalledOnce();
      expect(callTool).toHaveBeenCalledWith('alice', 'approval_response', {
        decisions: [{ approvalId: 'ap-1', decision: 'approve' }],
      });
      expect(send).toHaveBeenCalledOnce();
      expect(JSON.stringify(send.mock.calls[0][0])).toContain('resumed text');
```

Add a test for `groupChats: false`:

```ts
  it('rejects group approval actions when group chats are disabled', async () => {
    const { app, deps } = setup();
    const factory = vi.fn<TeamsAppFactory>(() => app);
    const dir = await mkdtemp(join(tmpdir(), 'teams-approval-disabled-'));
    const connector = new TeamsConnector(
      { stateDir: dir, transport: 'socket', groupChats: false }, deps, factory,
    );
    const callTool = vi.fn(async () => ({ result: { status: 'success' }, conversationId: 'c' }));
    await connector.start({ callTool } as never);
    try {
      const [entry, handlers] = factory.mock.calls[0];
      const result = await handlers.onApproval(entry, {
        send: vi.fn(),
        activity: {
          from: { aadObjectId: 'aad-1' },
          channelData: { tenant: { id: 't' } },
          conversation: { conversationType: 'groupChat' },
          value: { action: { data: { approvalId: 'ap-1' } } },
        },
      } as never, 'approve');
      expect(result).toHaveProperty('body.value', 'Only personal chats are supported.');
      expect(callTool).not.toHaveBeenCalled();
    } finally { await connector.stop(); }
  });
```

- [ ] **Step 5: Run the approval tests to verify they fail**

Run: `npx vitest run src/render.test.ts src/connector.test.ts`
Expected: FAIL because `onApproval` still rejects group chats.

- [ ] **Step 6: Implement group approvals and v2 install records**

In `src/connector.ts`, update `onInstall` to write v2 fields:

```ts
    this.state.conversations[`${bot.appId}:${participant.id}`] = {
      conversationId: activity.conversation.id,
      conversationType: 'personal',
      externalId: aadObjectId,
      tenantId: activity.channelData?.tenant?.id,
      serviceUrl: activity.serviceUrl,
      displayName: activity.from?.name,
      updatedAt: new Date().toISOString(),
    };
```

In `onApproval`, replace the conversation-type check:

```ts
    const conversationType = context.activity.conversation?.conversationType;
    if (conversationType !== 'personal' && !(conversationType === 'groupChat' && this.options.groupChats)) {
      return approvalResponse(this.options.groupChats
        ? 'Only personal chats and group chats are supported.'
        : 'Only personal chats are supported.');
    }
```

After resolving `participant`, compute the decider name and pass it to the card:

```ts
    const deciderName = participant.name ?? context.activity.from?.name ?? participant.id;
```

```ts
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
```

- [ ] **Step 7: Run the approval and connector tests**

Run: `npx vitest run src/render.test.ts src/connector.test.ts`
Expected: PASS.

- [ ] **Step 8: Commit**

```bash
git add src/render.ts src/connector.ts src/render.test.ts src/connector.test.ts
git commit -m "feat(approvals): allow group decisions and show the decider"
```

---

### Task 8: Outbound group routing

**Files:**
- Modify: `src/outbound.ts`
- Test: `src/outbound.test.ts`

**Interfaces:**
- Consumes: `ConversationRecord`, `TeamsState.legionConversationTargets` from Task 2; `teams-group` identity convention from Task 3.
- Produces:
  - `deliverMessage` prefers `legionConversationTargets[message.conversationId]`.
  - `teams-group` recipients send directly to their `externalId`.
  - Group records are created under `<appId>:<participantId>:<conversationId>`.
  - Personal creation uses `conversationType: 'personal'` and `externalId`.

- [ ] **Step 1: Update the existing outbound tests for v2 records and target routing**

`deliverMessage` now requires `message.conversationId`. Add `conversationId: 'legion-1'` to every message object passed to `deliverMessage` in `src/outbound.test.ts`, including the shared `message` in the first test.

In `src/outbound.test.ts`, update the personal creation assertion:

```ts
    expect(s.state.conversations['app:alice']).toEqual(expect.objectContaining({
      conversationId: 'conv-new',
      conversationType: 'personal',
      externalId: 'aad-1',
      tenantId: 'tenant-1',
    }));
    expect(s.state.legionConversationTargets['legion-1']).toBe('app:alice');
```

The first test's hand-written records are fine because the function only reads `conversationId` and `updatedAt`; keep them.

- [ ] **Step 2: Add failing group outbound tests**

Append to `src/outbound.test.ts`:

```ts
  it('routes a Legion conversation to its target group', async () => {
    const s = setup(
      [
        { id: 'assistant', identities: [{ connector: 'teams-bot', externalId: 'app' }] },
        { id: 'team-eng', identities: [{ connector: 'teams-group', externalId: '19:g@thread.v2' }] },
      ],
      {
        'app:team-eng:19:g@thread.v2': { conversationId: '19:g@thread.v2', updatedAt: 'now' },
        'app:team-eng:19:other@thread.v2': { conversationId: '19:other@thread.v2', updatedAt: 'later' },
      },
    );
    s.state.legionConversationTargets['legion-1'] = 'app:team-eng:19:g@thread.v2';
    await deliverMessage(
      { collective: s.collective, state: s.state as never, stateStore: { save: s.save } as never,
        apps: s.apps as never, options: resolveOptions({ stateDir: '/tmp/x' }) },
      { senderId: 'assistant', recipientId: 'team-eng', conversationId: 'legion-1', content: 'hello' },
    );
    expect(s.send).toHaveBeenCalledWith('19:g@thread.v2', expect.objectContaining({ text: 'hello' }));
    expect(s.send).not.toHaveBeenCalledWith('19:other@thread.v2', expect.anything());
  });

  it('sends directly to a teams-group identity and records the target', async () => {
    const s = setup([
      { id: 'assistant', identities: [{ connector: 'teams-bot', externalId: 'app' }] },
      { id: 'team-eng', identities: [{ connector: 'teams-group', externalId: '19:g@thread.v2' }] },
    ]);
    await deliverMessage(
      { collective: s.collective, state: s.state as never, stateStore: { save: s.save } as never,
        apps: s.apps as never, options: resolveOptions({ stateDir: '/tmp/x' }) },
      { senderId: 'assistant', recipientId: 'team-eng', conversationId: 'legion-1', content: 'hello' },
    );
    expect(s.send).toHaveBeenCalledWith('19:g@thread.v2', expect.objectContaining({ text: 'hello' }));
    expect(s.state.conversations['app:team-eng:19:g@thread.v2']).toMatchObject({
      conversationId: '19:g@thread.v2',
      conversationType: 'groupChat',
      externalId: '19:g@thread.v2',
    });
    expect(s.state.legionConversationTargets['legion-1']).toBe('app:team-eng:19:g@thread.v2');
    expect(s.save).toHaveBeenCalledWith(s.state);
    expect(s.create).not.toHaveBeenCalled();
  });

  it('keeps fallback group replies in the originating group', async () => {
    const s = setup(
      [
        { id: 'assistant', identities: [{ connector: 'teams-bot', externalId: 'app' }] },
        { id: 'operator', identities: [{ connector: 'teams-group', externalId: '19:g1@thread.v2' }] },
      ],
      {
        'app:operator:19:g1@thread.v2': { conversationId: '19:g1@thread.v2', updatedAt: 'now' },
        'app:operator:19:g2@thread.v2': { conversationId: '19:g2@thread.v2', updatedAt: 'later' },
      },
    );
    s.state.legionConversationTargets['legion-1'] = 'app:operator:19:g1@thread.v2';
    await deliverMessage(
      { collective: s.collective, state: s.state as never, stateStore: { save: s.save } as never,
        apps: s.apps as never, options: resolveOptions({ stateDir: '/tmp/x' }) },
      { senderId: 'assistant', recipientId: 'operator', conversationId: 'legion-1', content: 'hello' },
    );
    expect(s.send).toHaveBeenCalledWith('19:g1@thread.v2', expect.objectContaining({ text: 'hello' }));
    expect(s.send).not.toHaveBeenCalledWith('19:g2@thread.v2', expect.anything());
  });
```

- [ ] **Step 3: Run the outbound tests to verify they fail**

Run: `npx vitest run src/outbound.test.ts`
Expected: FAIL because group identities and the target map are not handled.

- [ ] **Step 4: Implement the outbound routing**

Replace `src/outbound.ts` with:

```ts
// src/outbound.ts
import type { CollectiveLike } from './directory.js';
import { MessageActivityInput } from '@microsoft/teams.api';
import type { ResolvedOptions } from './options.js';
import { chunkMarkdown } from './render.js';
import { describeError } from './errors.js';
import type { ConversationRecord, StateStore, TeamsState } from './state.js';

export interface OutboundApp {
  send(conversationId: string, activity: MessageActivityInput): Promise<unknown>;
  api: {
    conversations: {
      create(params: Record<string, unknown>): Promise<{ id: string }>;
    };
  };
}

export interface OutboundDeps {
  collective: CollectiveLike;
  state: TeamsState;
  stateStore: StateStore;
  apps: { getByAppId(appId: string): OutboundApp | undefined; ownerOf(appId: string): string | undefined };
  options: ResolvedOptions;
}

export async function deliverMessage(
  deps: OutboundDeps,
  message: { senderId: string; recipientId: string; conversationId: string; content: string },
): Promise<void> {
  const sender = deps.collective.get(message.senderId);
  const botIdentity = sender?.identities?.find((i) => i.connector === 'teams-bot');
  if (!botIdentity || deps.apps.ownerOf(botIdentity.externalId) !== message.senderId) return;

  const app = deps.apps.getByAppId(botIdentity.externalId);
  if (!app) return;

  const appId = botIdentity.externalId;
  const recipient = deps.collective.get(message.recipientId);
  let record: ConversationRecord | undefined;

  const targetKey = deps.state.legionConversationTargets[message.conversationId];
  if (targetKey) record = deps.state.conversations[targetKey];

  if (!record) {
    const groupIdentity = recipient?.identities?.find((i) => i.connector === 'teams-group');
    if (groupIdentity) {
      const key = `${appId}:${message.recipientId}:${groupIdentity.externalId}`;
      record = deps.state.conversations[key] ?? {
        conversationId: groupIdentity.externalId,
        conversationType: 'groupChat',
        externalId: groupIdentity.externalId,
        updatedAt: new Date().toISOString(),
      };
      deps.state.conversations[key] = record;
      deps.state.legionConversationTargets[message.conversationId] = key;
      await deps.stateStore.save(deps.state);
    } else {
      const key = `${appId}:${message.recipientId}`;
      record = deps.state.conversations[key];

      if (!record) {
        const userIdentity = recipient?.identities?.find((i) => i.connector === 'teams');
        if (userIdentity) {
          let conversation: { id: string };
          try {
            conversation = await app.api.conversations.create({
              members: [{ id: userIdentity.externalId }],
              tenantId: deps.options.tenantId,
              channelData: { tenant: { id: deps.options.tenantId } },
            });
          } catch (error) {
            console.error('Teams conversation creation failed', describeError(error));
            return;
          }
          record = {
            conversationId: conversation.id,
            conversationType: 'personal',
            externalId: userIdentity.externalId,
            tenantId: deps.options.tenantId,
            updatedAt: new Date().toISOString(),
          };
          deps.state.conversations[key] = record;
          deps.state.legionConversationTargets[message.conversationId] = key;
          await deps.stateStore.save(deps.state);
        } else {
          for (const [candidateKey, candidate] of Object.entries(deps.state.conversations)) {
            if (candidateKey.startsWith(`${key}:`) && (!record || candidate.updatedAt > record.updatedAt)) {
              record = candidate;
            }
          }
        }
      }
    }
  }
  if (!record) return;

  for (const chunk of chunkMarkdown(message.content, deps.options.maxMessageChars)) {
    const activity = new MessageActivityInput(chunk);
    if (deps.options.aiLabel) activity.addAiGenerated();
    try {
      await app.send(record.conversationId, activity);
    } catch (error) {
      console.error('Teams message delivery failed', describeError(error));
      return;
    }
  }
}
```

- [ ] **Step 5: Run the outbound tests**

Run: `npx vitest run src/outbound.test.ts`
Expected: PASS.

- [ ] **Step 6: Commit**

```bash
git add src/outbound.ts src/outbound.test.ts
git commit -m "feat(outbound): route replies to the originating group"
```

---

### Task 9: Documentation and full verification

**Files:**
- Modify: `README.md`
- Modify: `docs/deploy.md`

- [ ] **Step 1: Update README**

Change the opening paragraph to say the connector exposes participants in personal 1:1 chats and group chats, and that team channels remain unsupported.

Add to the options table:

```markdown
| `groupChats` | `true` | Enable group-chat participation. |
| `groupSenderPrefix` | `"[{name}]: {message}"` | Sender prefix for group messages; `null` disables it and the template must contain `{message}`. |
```

Add a group-chat paragraph after the personal-user identity paragraph:

```markdown
Group chats are modelled as one Legion `user` participant with a `teams-group` identity whose `externalId` is the Teams conversation ID. The group uses the same `unknownUserPolicy` as personal chats. With `auto-provision`, the participant ID is `teams-group-<name>-<hash>`; unnamed groups use `Teams group <hash>`. Group messages are prefixed with the sender's name by default. `/new` and `/whoami` act on the group, while `/pair` runs as the registered human who typed it. Replies are routed back through the Legion conversation target map, so a fallback participant paired to several groups still answers in the originating group.
```

Update the streaming note:

```markdown
Teams streaming is supported in 1:1 and group chats; team channels are not supported. Teams allows one open stream per chat, so inbound turns are serialized per conversation.
```

Update the fallback paragraph to mention group conversation isolation and the target map.

Update the commands section:

```markdown
- `/new` — start a new Legion conversation for the current chat (the group, in a group chat).
- `/pair <code> <participant-id>` — pair a pending identity; in a group chat it runs as the registered human who typed it.
- `/whoami` — show the current participant, chat, sender, and bot app.
- `/help` — list commands.
```

- [ ] **Step 2: Update deployment notes**

In `docs/deploy.md`, replace the sentence that says the connector only accepts personal chats:

```markdown
The connector accepts personal chats and group chats, and checks that incoming messages, install events, and approval actions match the tenant in the bot's credentials. Team channels are not supported yet. With shared fallback enabled, inbound personal threads are isolated per Teams sender and inbound group threads are isolated per conversation; proactive routing uses the Legion conversation target map, then an exact participant route, then the last active fallback chat.
```

- [ ] **Step 3: Run the full test suite**

Run: `npm test`
Expected: all tests PASS; the live integration test remains skipped unless `LEGION_TEAMS_INTEGRATION=1`.

- [ ] **Step 4: Typecheck and build**

Run: `npm run typecheck`
Expected: PASS.

Run: `npm run build`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add README.md docs/deploy.md
git commit -m "docs: document Teams group-chat support"
```

- [ ] **Step 6: Live verification (manual, after unit verification)**

1. Start Legion with the connector in the local live workspace (Socket Mode, bot **Assistant**).
2. Create or open a Teams group chat with the bot installed and @mention it.
3. Verify:
   - The bot replies in the group.
   - The agent sees `[<sender name>]: <message>`.
   - Streaming appears in the group.
   - A second @mention while the first turn is running waits for the first to finish.
   - A bare @mention sends an empty message and the agent responds.
   - An approval card posted in the group can be decided by a registered member; the decided card shows the outcome, optional reason, and the decider's name.
   - `/new` clears the group conversation; `/whoami` shows the group participant and conversation ID; `/pair` runs as the typing human.
4. Confirm the pending personal-chat retests from the previous round: rejection reason visibility, decided-card reason display, and typing indicator stopping promptly.

---

## Self-Review

**Spec coverage**

| Spec section | Task |
| --- | --- |
| Group identity `teams-group` / conversation ID | 3, 5 |
| Same `unknownUserPolicy` | 3, 5 |
| Auto-provision ID and naming | 3, 5 |
| Mention stripping (bot only) | 5 |
| Configurable attribution | 1, 5 |
| Group conversation records | 2, 5 |
| Group installs ignored | 2, 5, 7 (install test) |
| Pairing codes generalized | 2, 5 |
| Retired groups treated as unknown | 3 |
| `/new`, `/whoami`, `/pair` | 4, 5 |
| Native streaming and typing | 5 (unchanged `communicate` path) |
| One stream per chat | 6 |
| Group approvals and decider | 7 |
| Outbound target map and group routing | 8 |
| State v2 and migration | 2 |
| Options | 1 |
| README/deploy docs | 9 |

**Placeholder scan:** every code step contains concrete code; every test step names the command and expected result. No TBD/TODO entries.

**Type consistency:** `PendingPairKey`, `ConversationRecord`, `ResolvedGroup`, `GroupIdentityOutcome`, `CommandDeps`, `SerialQueue`, and `DecidedApprovalView` are introduced once and consumed with the same field names in later tasks. The state key formats are consistent: personal `<appId>:<participantId>`, fallback personal `<appId>:<participantId>:<externalId>`, group `<appId>:<participantId>:<conversationId>`. The queue key is the platform chat identity (`<appId>:<group conversationId>` or `<appId>:<sender aadObjectId>`), which is stable for a whole turn and one per chat.
