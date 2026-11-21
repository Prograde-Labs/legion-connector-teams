# Teams Connector Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a Microsoft Teams connector to Legion (one bot per participant, personal 1:1 chats) plus the connector-agnostic core `pair` tool it depends on.

**Architecture:** A new `@legion-collective/connector-teams` module implements the Legion `Connector` contract. It scans the collective for `teams-bot` identities, runs one Teams SDK `App` per bot, resolves inbound `aadObjectId`s to Legion users (pairing / auto-provision / reject / fallback), streams `communicate` output, renders approvals as Adaptive Cards, and delivers outbound messages through the sender's bot. A small core `pair` tool writes identities into participant files so AuthEngine can gate pairing for any participant, agent, or service.

**Tech Stack:** TypeScript ESM/NodeNext, Node 22, `@microsoft/teams.apps` 2.2.1, `@microsoft/teams.api` 2.2.1, `@microsoft/teams.cards` 2.2.1, Express 5, Vitest.

**Spec:** `docs/specs/2026-11-21-teams-connector-design.md`

## Global Constraints

- Node **22.x** (`node --version`). Legion's Docker image is `node:22-slim`.
- Teams SDK packages pinned to **2.2.1** (`apps`, `api`, `cards`).
- Legion packages are dev-only `file:../legion/packages/core` and `file:../legion/packages/types`.
- Connector name `teams`; bot identity connector `teams-bot`; user identity connector `teams`.
- One Teams bot per participant; personal 1:1 scope only in this plan.
- HTTP transport is the default; Socket Mode is opt-in (`transport: 'socket'`).
- No secrets in git. Bot credentials live only in the Legion SecretStore under `teams-bot/<appId>`.
- Every task ends with tests passing and a commit.
- Connector test command: `npm test` (Vitest). Typecheck: `npm run typecheck`.
- Legion test command: `npx vitest run <path>` from `repos/legion`.

## File Structure

**Legion core (`repos/legion`)**
- Create `packages/core/src/tools/pair-tool.ts` — connector-agnostic `pair` tool.
- Create `packages/core/src/tools/pair-tool.test.ts`.
- Modify `packages/core/src/tools/user-tools.ts` — export `pairTool` in `userTools`.
- Modify `packages/core/src/collective/default-participants.ts` — add `pair` to `MANAGEMENT_TOOLS`.

**Connector (`repos/legion-connector-teams`)**
- Create `src/options.ts`, `src/connector.ts`, `src/index.ts`.
- Create `src/state.ts`, `src/secrets.ts`, `src/directory.ts`.
- Create `src/pairing.ts`, `src/identity.ts`, `src/commands.ts`.
- Create `src/render.ts`, `src/communicate.ts`, `src/approvals.ts`.
- Create `src/inbound.ts`, `src/outbound.ts`, `src/bot-pool.ts`.
- Create `src/test-helpers.ts`; tests beside each module.
- Create `scripts/provision-teams-bot.mjs`, `scripts/provision-all.mjs`.
- Create `README.md`, `docs/deploy.md`.

---

## Task 0: Workspace setup

**Files:** none (environment only).

- [ ] **Step 1: Create the Legion worktree for the core change**

```bash
cd /home/chris/workspace
mkdir -p threads/2026-11-21-teams-connector
git -C repos/legion worktree add threads/2026-11-21-teams-connector/legion -b feat/teams-pair-tool main
```

- [ ] **Step 2: Build Legion so the connector can resolve core types**

```bash
cd /home/chris/workspace/repos/legion && npm run build
```

- [ ] **Step 3: Install connector dependencies**

```bash
cd /home/chris/workspace/repos/legion-connector-teams && npm install
```

Expected: `node_modules/` and `package-lock.json` created; no errors.

- [ ] **Step 4: Verify Node version**

Run: `node --version` — Expected: `v22.x`.

---

## Task 1: Core `pair` tool

**Files:**
- Create: `packages/core/src/tools/pair-tool.ts`
- Create: `packages/core/src/tools/pair-tool.test.ts`
- Modify: `packages/core/src/tools/user-tools.ts`
- Modify: `packages/core/src/collective/default-participants.ts`

**Interfaces:**
- Produces: `pairTool: Tool` with `name = 'pair'`, parameters `{ participantId, connector, externalId }`.
- Consumes: `ToolContext.collective` (`get`, `findByIdentity`, `update`).

- [ ] **Step 1: Write the failing test**

```ts
// packages/core/src/tools/pair-tool.test.ts
import { describe, expect, it } from 'vitest';
import { pairTool } from './pair-tool.js';
import type { ToolContext } from './Tool.js';
import type { ParticipantConfig } from '@legion-collective/types';

function makeContext(participants: Record<string, ParticipantConfig>) {
  const updates: Array<{ id: string; patch: Partial<ParticipantConfig> }> = [];
  const collective = {
    get: (id: string) => participants[id],
    findByIdentity: (connector: string, externalId: string) =>
      Object.values(participants).find((p) =>
        p.identities?.some((i) => i.connector === connector && i.externalId === externalId),
      ),
    update: async (id: string, patch: Partial<ParticipantConfig>) => {
      updates.push({ id, patch });
      participants[id] = { ...participants[id], ...patch } as ParticipantConfig;
    },
  };
  return { context: { collective } as unknown as ToolContext, updates, participants };
}

const user = (id: string, identities: ParticipantConfig['identities'] = []): ParticipantConfig => ({
  id,
  name: id,
  type: 'user',
  tools: {},
  identities,
});

describe('pair tool', () => {
  it('merges an identity into a participant', async () => {
    const { context, participants } = makeContext({ alice: user('alice') });
    const result = await pairTool.execute(
      { participantId: 'alice', connector: 'teams', externalId: 'aad-1' },
      context,
    );
    expect(result).toMatchObject({ status: 'success' });
    expect(participants.alice.identities).toEqual([
      { connector: 'teams', externalId: 'aad-1' },
    ]);
  });

  it('dedupes an existing identity', async () => {
    const { context } = makeContext({
      alice: user('alice', [{ connector: 'teams', externalId: 'aad-1' }]),
    });
    const result = await pairTool.execute(
      { participantId: 'alice', connector: 'teams', externalId: 'aad-1' },
      context,
    );
    expect(result).toMatchObject({ status: 'success', data: { alreadyPaired: true } });
  });

  it('preserves unrelated identities', async () => {
    const { context, participants } = makeContext({
      alice: user('alice', [{ connector: 'web', externalId: 'alice' }]),
    });
    await pairTool.execute(
      { participantId: 'alice', connector: 'teams', externalId: 'aad-1' },
      context,
    );
    expect(participants.alice.identities).toEqual([
      { connector: 'web', externalId: 'alice' },
      { connector: 'teams', externalId: 'aad-1' },
    ]);
  });

  it('rejects a missing participant', async () => {
    const { context } = makeContext({});
    const result = await pairTool.execute(
      { participantId: 'ghost', connector: 'teams', externalId: 'aad-1' },
      context,
    );
    expect(result).toMatchObject({ status: 'error' });
  });

  it('rejects empty fields', async () => {
    const { context } = makeContext({ alice: user('alice') });
    const result = await pairTool.execute(
      { participantId: 'alice', connector: '', externalId: 'aad-1' },
      context,
    );
    expect(result).toMatchObject({ status: 'error' });
  });

  it('rejects an identity owned by another participant', async () => {
    const { context } = makeContext({
      alice: user('alice'),
      bob: user('bob', [{ connector: 'teams', externalId: 'aad-1' }]),
    });
    const result = await pairTool.execute(
      { participantId: 'alice', connector: 'teams', externalId: 'aad-1' },
      context,
    );
    expect(result).toMatchObject({ status: 'error' });
    expect((result as { error: string }).error).toMatch(/already paired to bob/);
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run packages/core/src/tools/pair-tool.test.ts`
Expected: FAIL — cannot find `./pair-tool.js`.

- [ ] **Step 3: Write the implementation**

```ts
// packages/core/src/tools/pair-tool.ts
import type { ConnectorIdentity, ToolResult } from '@legion-collective/types';
import type { Tool, ToolContext } from './Tool.js';

function requireCollective(context: ToolContext): NonNullable<ToolContext['collective']> {
  if (!context.collective) throw new Error('collective unavailable in context');
  return context.collective;
}

export const pairTool: Tool = {
  name: 'pair',
  description: 'Merge an external identity into a participant (connector-agnostic).',
  parameters: {
    type: 'object',
    properties: {
      participantId: { type: 'string', description: 'Target participant ID' },
      connector: { type: 'string', description: 'Connector name, e.g. teams or telegram' },
      externalId: { type: 'string', description: 'Platform identity' },
    },
    required: ['participantId', 'connector', 'externalId'],
  },
  async execute(args, context): Promise<ToolResult> {
    const { participantId, connector, externalId } = (args ?? {}) as {
      participantId?: string;
      connector?: string;
      externalId?: string;
    };
    if (!participantId || !connector || !externalId) {
      return { status: 'error', error: 'participantId, connector, and externalId are required' };
    }
    const collective = requireCollective(context);
    const target = collective.get(participantId);
    if (!target) return { status: 'error', error: `Participant not found: ${participantId}` };

    const owner = collective.findByIdentity(connector, externalId);
    if (owner && owner.id !== participantId) {
      return {
        status: 'error',
        error: `Identity ${connector}/${externalId} is already paired to ${owner.id}`,
      };
    }

    const identities = target.identities ?? [];
    if (identities.some((i) => i.connector === connector && i.externalId === externalId)) {
      return {
        status: 'success',
        data: { participantId, connector, externalId, alreadyPaired: true },
      };
    }

    const next: ConnectorIdentity[] = [...identities, { connector, externalId }];
    await collective.update(participantId, { identities: next });
    return { status: 'success', data: { participantId, connector, externalId } };
  },
};
```

- [ ] **Step 4: Register the tool**

In `packages/core/src/tools/user-tools.ts`, add:

```ts
import { pairTool } from './pair-tool.js';
```

and change the export to:

```ts
export const userTools: Tool[] = [createUserTool, modifyUserTool, pairTool];
```

In `packages/core/src/collective/default-participants.ts`, add `'pair'` to `MANAGEMENT_TOOLS` (next to `'create_user'` and `'modify_user'`).

- [ ] **Step 5: Run tests**

Run: `npx vitest run packages/core/src/tools/pair-tool.test.ts`
Expected: PASS (6 tests).

Run: `npx vitest run packages/core/src/tools/user-tools.test.ts packages/core/src/collective/default-participants.test.ts`
Expected: PASS.

- [ ] **Step 6: Commit**

```bash
cd /home/chris/workspace/threads/2026-11-21-teams-connector/legion
git add packages/core/src/tools/pair-tool.ts packages/core/src/tools/pair-tool.test.ts packages/core/src/tools/user-tools.ts packages/core/src/collective/default-participants.ts
git commit -m "feat(core): add connector-agnostic pair tool"
```

---

## Task 2: Connector scaffolding and factory

**Files:**
- Create: `package.json`, `tsconfig.json`, `vitest.config.ts`, `.gitignore`
- Create: `src/options.ts`, `src/connector.ts`, `src/index.ts`
- Create: `src/index.test.ts`

**Interfaces:**
- Produces: `connector(options, deps): Connector` with `name = 'teams'`.
- Produces: `TeamsConnectorOptions`, `ResolvedOptions`, `resolveOptions()`.

- [ ] **Step 1: Create package metadata**

```json
{
  "name": "@legion-collective/connector-teams",
  "version": "0.1.0",
  "description": "Microsoft Teams boundary connector for Legion collectives",
  "type": "module",
  "main": "dist/index.js",
  "types": "dist/index.d.ts",
  "files": ["dist"],
  "scripts": {
    "build": "tsc --project tsconfig.json",
    "test": "vitest run",
    "test:watch": "vitest",
    "typecheck": "tsc --noEmit"
  },
  "engines": { "node": ">=22" },
  "dependencies": {
    "@microsoft/teams.apps": "2.2.1",
    "@microsoft/teams.api": "2.2.1",
    "@microsoft/teams.cards": "2.2.1",
    "express": "^5.3.0"
  },
  "devDependencies": {
    "@legion-collective/core": "file:../legion/packages/core",
    "@legion-collective/types": "file:../legion/packages/types",
    "@types/express": "^5.0.0",
    "@types/node": "^22.0.0",
    "typescript": "^5.6.0",
    "vitest": "^2.1.0"
  }
}
```

```json
// tsconfig.json
{
  "compilerOptions": {
    "target": "ES2022",
    "module": "NodeNext",
    "moduleResolution": "NodeNext",
    "lib": ["ES2022"],
    "outDir": "dist",
    "rootDir": "src",
    "strict": true,
    "esModuleInterop": true,
    "noUnusedLocals": true,
    "noUnusedParameters": true,
    "verbatimModuleSyntax": true,
    "declaration": true,
    "skipLibCheck": true,
    "types": ["node"]
  },
  "include": ["src"]
}
```

```ts
// vitest.config.ts
import { defineConfig } from 'vitest/config';
export default defineConfig({ test: { include: ['src/**/*.test.ts'] } });
```

```text
# .gitignore
node_modules/
dist/
*.tsbuildinfo
.env
*.zip
```

- [ ] **Step 2: Write the failing test**

```ts
// src/index.test.ts
import { describe, expect, it } from 'vitest';
import { connector } from './index.js';
import type { ConnectorRuntimeDeps } from '@legion-collective/core';

const deps = (overrides: Partial<ConnectorRuntimeDeps> = {}): ConnectorRuntimeDeps =>
  ({ collective: {} as never, secrets: {} as never, ...overrides });

describe('connector factory', () => {
  it('returns a connector named teams', () => {
    expect(connector({}, deps()).name).toBe('teams');
  });

  it('throws at start without a secret store', async () => {
    const c = connector({}, deps({ secrets: undefined }));
    await expect(c.start({} as never)).rejects.toThrow(/LEGION_SECRETS_KEY/);
  });

  it('rejects fallback policy without fallbackUserId', async () => {
    const c = connector({ unknownUserPolicy: 'fallback' }, deps());
    await expect(c.start({} as never)).rejects.toThrow(/fallbackUserId/);
  });
});
```

- [ ] **Step 3: Run test to verify it fails**

Run: `npm test -- src/index.test.ts`
Expected: FAIL — `src/index.ts` missing.

- [ ] **Step 4: Implement options, connector, and factory**

```ts
// src/options.ts
import type { DetailLevel } from '@legion-collective/types';

export type TeamsTransport = 'http' | 'socket';
export type UnknownUserPolicy = 'pairing' | 'reject' | 'auto-provision' | 'fallback';
export type ToolDisplay = 'off' | 'status' | 'cards';

export interface TeamsConnectorOptions {
  transport?: TeamsTransport;
  port?: number;
  publicBaseUrl?: string;
  tenantId?: string;
  stateDir?: string;
  rescanIntervalMs?: number;
  socketGeos?: string[];
  unknownUserPolicy?: UnknownUserPolicy;
  fallbackUserId?: string;
  provisionerId?: string;
  pairingParticipantId?: string;
  pairingTtlSeconds?: number;
  maxDetail?: DetailLevel;
  toolDisplay?: ToolDisplay;
  aiLabel?: boolean;
  feedback?: boolean;
  typingIntervalMs?: number;
  maxMessageChars?: number;
}

export interface ResolvedOptions {
  transport: TeamsTransport;
  port: number;
  publicBaseUrl?: string;
  tenantId?: string;
  stateDir: string;
  rescanIntervalMs: number;
  socketGeos: string[];
  unknownUserPolicy: UnknownUserPolicy;
  fallbackUserId?: string;
  provisionerId: string;
  pairingParticipantId?: string;
  pairingTtlSeconds: number;
  maxDetail: DetailLevel;
  toolDisplay: ToolDisplay;
  aiLabel: boolean;
  feedback: boolean;
  typingIntervalMs: number;
  maxMessageChars: number;
}

export function resolveOptions(options: TeamsConnectorOptions): ResolvedOptions {
  const workspace = process.env.LEGION_WORKSPACE ?? process.cwd();
  return {
    transport: options.transport ?? 'http',
    port: options.port ?? 3978,
    publicBaseUrl: options.publicBaseUrl,
    tenantId: options.tenantId,
    stateDir: options.stateDir ?? `${workspace}/.legion/connectors/teams`,
    rescanIntervalMs: options.rescanIntervalMs ?? 60_000,
    socketGeos: options.socketGeos ?? ['amer', 'emea', 'apac'],
    unknownUserPolicy: options.unknownUserPolicy ?? 'pairing',
    fallbackUserId: options.fallbackUserId,
    provisionerId: options.provisionerId ?? 'operator',
    pairingParticipantId: options.pairingParticipantId,
    pairingTtlSeconds: options.pairingTtlSeconds ?? 600,
    maxDetail: options.maxDetail ?? 'standard',
    toolDisplay: options.toolDisplay ?? 'status',
    aiLabel: options.aiLabel ?? true,
    feedback: options.feedback ?? false,
    typingIntervalMs: options.typingIntervalMs ?? 4000,
    maxMessageChars: options.maxMessageChars ?? 20000,
  };
}
```

```ts
// src/connector.ts (minimal; expanded in Task 12)
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
    this.ctx = undefined;
  }
}
```

```ts
// src/index.ts
import type { Connector, ConnectorRuntimeDeps } from '@legion-collective/core';
import { TeamsConnector } from './connector.js';
import type { TeamsConnectorOptions } from './options.js';

export type { TeamsConnectorOptions } from './options.js';

export function connector(
  options: TeamsConnectorOptions,
  deps: ConnectorRuntimeDeps,
): Connector {
  return new TeamsConnector(options, deps);
}
```

- [ ] **Step 5: Run tests**

Run: `npm test -- src/index.test.ts`
Expected: PASS (3 tests).

- [ ] **Step 6: Commit**

```bash
cd /home/chris/workspace/repos/legion-connector-teams
git add package.json tsconfig.json vitest.config.ts .gitignore src/options.ts src/connector.ts src/index.ts src/index.test.ts
git commit -m "feat: scaffold Teams connector package"
```

---

## Task 3: Persistent state store

**Files:**
- Create: `src/state.ts`
- Create: `src/state.test.ts`

**Interfaces:**
- Produces: `TeamsState`, `ConversationRecord`, `PendingPair`, `emptyState()`, `StateStore` with `load()` / `save()`.

- [ ] **Step 1: Write the failing test**

```ts
// src/state.test.ts
import { mkdtemp, readFile, stat, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { emptyState, StateStore } from './state.js';

describe('StateStore', () => {
  it('returns empty state when the file is missing', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'teams-state-'));
    await expect(new StateStore(dir).load()).resolves.toEqual(emptyState());
  });

  it('round-trips state', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'teams-state-'));
    const store = new StateStore(dir);
    const state = emptyState();
    state.conversations['app:alice'] = {
      conversationId: 'conv-1',
      updatedAt: '2026-11-21T00:00:00.000Z',
    };
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

  it('backs up corrupt state and starts fresh', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'teams-state-'));
    await writeFile(join(dir, 'state.json'), '{not json');
    const store = new StateStore(dir);
    await expect(store.load()).resolves.toEqual(emptyState());
    const raw = await readFile(join(dir, 'state.json'), 'utf8').catch(() => null);
    expect(raw).toBeNull();
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npm test -- src/state.test.ts`
Expected: FAIL — `./state.js` missing.

- [ ] **Step 3: Implement**

```ts
// src/state.ts
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
```

- [ ] **Step 4: Run tests**

Run: `npm test -- src/state.test.ts`
Expected: PASS (4 tests).

- [ ] **Step 5: Commit**

```bash
git add src/state.ts src/state.test.ts
git commit -m "feat: add persistent Teams connector state"
```

---

## Task 4: Bot secrets and directory scan

**Files:**
- Create: `src/secrets.ts`, `src/directory.ts`
- Create: `src/secrets.test.ts`, `src/directory.test.ts`

**Interfaces:**
- Produces: `BotCredentials`, `parseBotSecret(raw)`, `resolveBotCredentials(secrets, participantId, appId)`.
- Produces: `BotEntry { participantId, name, appId, credentials }`, `DirectoryResult { bots, byParticipant, warnings }`, `buildDirectory(collective, secrets)`.
- Consumes: `CollectiveLike.list()`, `SecretStoreLike.getSecret()`.

- [ ] **Step 1: Write the failing tests**

```ts
// src/secrets.test.ts
import { describe, expect, it } from 'vitest';
import { parseBotSecret, resolveBotCredentials } from './secrets.js';

describe('bot secrets', () => {
  it('parses a valid credential JSON', () => {
    expect(parseBotSecret('{"appId":"a","clientSecret":"s","tenantId":"t"}'))
      .toEqual({ appId: 'a', clientSecret: 's', tenantId: 't' });
  });

  it('returns null for missing or malformed values', () => {
    expect(parseBotSecret(null)).toBeNull();
    expect(parseBotSecret('not-json')).toBeNull();
    expect(parseBotSecret('{"appId":"a"}')).toBeNull();
  });

  it('reads the participant-scoped secret key', async () => {
    const secrets = {
      getSecret: async (participantId: string, key: string) =>
        participantId === 'assistant' && key === 'teams-bot/app-1'
          ? '{"appId":"app-1","clientSecret":"s","tenantId":"t"}'
          : null,
    };
    await expect(resolveBotCredentials(secrets, 'assistant', 'app-1'))
      .resolves.toEqual({ appId: 'app-1', clientSecret: 's', tenantId: 't' });
    await expect(resolveBotCredentials(secrets, 'assistant', 'app-2')).resolves.toBeNull();
  });
});
```

```ts
// src/directory.test.ts
import { describe, expect, it } from 'vitest';
import { buildDirectory } from './directory.js';
import type { CollectiveLike, SecretStoreLike } from '@legion-collective/core';

const collective = (participants: Array<Record<string, unknown>>) =>
  ({ list: () => participants }) as unknown as CollectiveLike;

const secrets = (values: Record<string, string>): SecretStoreLike =>
  ({ getSecret: async (id: string, key: string) => values[`${id}:${key}`] ?? null }) as SecretStoreLike;

describe('buildDirectory', () => {
  it('includes only participants with resolvable bot credentials', async () => {
    const result = await buildDirectory(
      collective([
        { id: 'assistant', name: 'Assistant', type: 'agent', tools: {}, status: 'active',
          identities: [{ connector: 'teams-bot', externalId: 'app-1' }] },
        { id: 'researcher', name: 'Researcher', type: 'agent', tools: {}, status: 'active',
          identities: [{ connector: 'teams-bot', externalId: 'app-2' }] },
        { id: 'operator', name: 'Operator', type: 'user', tools: {}, status: 'active',
          identities: [{ connector: 'teams', externalId: 'aad-op' }] },
        { id: 'retired', name: 'Retired', type: 'agent', tools: {}, status: 'retired',
          identities: [{ connector: 'teams-bot', externalId: 'app-3' }] },
      ]),
      secrets({ 'assistant:teams-bot/app-1': '{"appId":"app-1","clientSecret":"s","tenantId":"t"}' }),
    );
    expect([...result.bots.keys()]).toEqual(['app-1']);
    expect(result.byParticipant.get('assistant')?.name).toBe('Assistant');
    expect(result.warnings.join(' ')).toMatch(/app-2/);
  });

  it('keeps the first bot on duplicate appId', async () => {
    const result = await buildDirectory(
      collective([
        { id: 'a', name: 'A', type: 'agent', tools: {},
          identities: [{ connector: 'teams-bot', externalId: 'dup' }] },
        { id: 'b', name: 'B', type: 'agent', tools: {},
          identities: [{ connector: 'teams-bot', externalId: 'dup' }] },
      ]),
      secrets({
        'a:teams-bot/dup': '{"appId":"dup","clientSecret":"s","tenantId":"t"}',
        'b:teams-bot/dup': '{"appId":"dup","clientSecret":"s","tenantId":"t"}',
      }),
    );
    expect(result.bots.get('dup')?.participantId).toBe('a');
  });
});
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `npm test -- src/secrets.test.ts src/directory.test.ts`
Expected: FAIL — modules missing.

- [ ] **Step 3: Implement**

```ts
// src/secrets.ts
import type { SecretStoreLike } from '@legion-collective/core';

export interface BotCredentials {
  appId: string;
  clientSecret: string;
  tenantId: string;
}

export function parseBotSecret(raw: string | null): BotCredentials | null {
  if (!raw) return null;
  try {
    const value = JSON.parse(raw) as Record<string, unknown>;
    if (
      typeof value.appId !== 'string' ||
      typeof value.clientSecret !== 'string' ||
      typeof value.tenantId !== 'string'
    ) {
      return null;
    }
    return {
      appId: value.appId,
      clientSecret: value.clientSecret,
      tenantId: value.tenantId,
    };
  } catch {
    return null;
  }
}

export async function resolveBotCredentials(
  secrets: SecretStoreLike,
  participantId: string,
  appId: string,
): Promise<BotCredentials | null> {
  return parseBotSecret(await secrets.getSecret(participantId, `teams-bot/${appId}`));
}
```

```ts
// src/directory.ts
import type { CollectiveLike, SecretStoreLike } from '@legion-collective/core';
import { resolveBotCredentials, type BotCredentials } from './secrets.js';

export interface BotEntry {
  participantId: string;
  name: string;
  appId: string;
  credentials: BotCredentials;
}

export interface DirectoryResult {
  bots: Map<string, BotEntry>;
  byParticipant: Map<string, BotEntry>;
  warnings: string[];
}

export async function buildDirectory(
  collective: CollectiveLike,
  secrets: SecretStoreLike,
): Promise<DirectoryResult> {
  const bots = new Map<string, BotEntry>();
  const byParticipant = new Map<string, BotEntry>();
  const warnings: string[] = [];

  for (const participant of collective.list()) {
    if ((participant.status ?? 'active') === 'retired') continue;
    const identity = participant.identities?.find((i) => i.connector === 'teams-bot');
    if (!identity) continue;

    const appId = identity.externalId;
    const existing = bots.get(appId);
    if (existing) {
      warnings.push(`duplicate teams-bot appId ${appId}; keeping ${existing.participantId}`);
      continue;
    }

    const credentials = await resolveBotCredentials(secrets, participant.id, appId);
    if (!credentials) {
      warnings.push(`missing or malformed teams-bot secret for ${participant.id} (${appId})`);
      continue;
    }
    if (credentials.appId !== appId) {
      warnings.push(`teams-bot secret appId mismatch for ${participant.id} (${appId})`);
      continue;
    }

    const entry: BotEntry = {
      participantId: participant.id,
      name: participant.name,
      appId,
      credentials,
    };
    bots.set(appId, entry);
    byParticipant.set(participant.id, entry);
  }

  return { bots, byParticipant, warnings };
}
```

- [ ] **Step 4: Run tests**

Run: `npm test -- src/secrets.test.ts src/directory.test.ts`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/secrets.ts src/directory.ts src/secrets.test.ts src/directory.test.ts
git commit -m "feat: resolve Teams bot credentials and directory"
```

---

## Task 5: Pairing codes and identity resolution

**Files:**
- Create: `src/pairing.ts`, `src/identity.ts`
- Create: `src/pairing.test.ts`, `src/identity.test.ts`

**Interfaces:**
- Produces: `generatePairCode(random?)`, `pruneExpiredPairs(state, now?)`, `rememberPendingPair(...)`, `consumePendingPair(state, code, now?)`.
- Produces: `ResolvedUser`, `IdentityOutcome`, `resolveIdentity({ collective, state, appId, aadObjectId, displayName, options, now })`.

- [ ] **Step 1: Write the failing tests**

```ts
// src/pairing.test.ts
import { describe, expect, it } from 'vitest';
import { consumePendingPair, generatePairCode, rememberPendingPair } from './pairing.js';
import { emptyState } from './state.js';

describe('pairing codes', () => {
  it('generates six unambiguous characters', () => {
    for (let i = 0; i < 20; i++) {
      expect(generatePairCode(() => 0)).toMatch(/^[2-9A-HJ-NP-Z]{6}$/);
    }
  });

  it('reuses an unexpired code for the same pair', () => {
    const state = emptyState();
    const first = rememberPendingPair(state, 'app', 'aad', 600, 1000);
    const second = rememberPendingPair(state, 'app', 'aad', 600, 2000);
    expect(second).toBe(first);
  });

  it('consumes a code once', () => {
    const state = emptyState();
    const code = rememberPendingPair(state, 'app', 'aad', 600, 1000);
    expect(consumePendingPair(state, code, 2000)).toMatchObject({ appId: 'app', aadObjectId: 'aad' });
    expect(consumePendingPair(state, code, 2000)).toBeNull();
  });

  it('rejects an expired code', () => {
    const state = emptyState();
    const code = rememberPendingPair(state, 'app', 'aad', 10, 1000);
    expect(consumePendingPair(state, code, 12_000)).toBeNull();
  });
});
```

```ts
// src/identity.test.ts
import { describe, expect, it } from 'vitest';
import { resolveIdentity } from './identity.js';
import { emptyState } from './state.js';
import { resolveOptions } from './options.js';
import type { CollectiveLike } from '@legion-collective/core';

const collective = (participants: Array<Record<string, unknown>>) =>
  ({
    list: () => participants,
    get: (id: string) => participants.find((p) => p.id === id),
    findByIdentity: (connector: string, externalId: string) =>
      participants.find((p) =>
        (p.identities as Array<{ connector: string; externalId: string }> | undefined)?.some(
          (i) => i.connector === connector && i.externalId === externalId,
        ),
      ),
  }) as unknown as CollectiveLike;

const options = (overrides = {}) => resolveOptions({ stateDir: '/tmp/x', ...overrides });

describe('resolveIdentity', () => {
  it('resolves an explicit teams identity', () => {
    const outcome = resolveIdentity({
      collective: collective([
        { id: 'alice', name: 'Alice', type: 'user', tools: {}, status: 'active',
          identities: [{ connector: 'teams', externalId: 'aad-1' }] },
      ]),
      state: emptyState(), appId: 'app', aadObjectId: 'aad-1', options: options(),
    });
    expect(outcome).toEqual({
      kind: 'resolved',
      user: { participantId: 'alice', aadObjectId: 'aad-1', displayName: undefined },
    });
  });

  it('issues a pairing code for an unknown user', () => {
    const state = emptyState();
    const outcome = resolveIdentity({
      collective: collective([]), state, appId: 'app', aadObjectId: 'aad-2', options: options(),
    });
    expect(outcome.kind).toBe('pairing');
    expect(Object.keys(state.pendingPairs)).toHaveLength(1);
  });

  it('returns auto-provision for that policy', () => {
    const outcome = resolveIdentity({
      collective: collective([]), state: emptyState(), appId: 'app', aadObjectId: 'aad-3',
      displayName: 'New User', options: options({ unknownUserPolicy: 'auto-provision' }),
    });
    expect(outcome).toEqual({ kind: 'auto-provision', aadObjectId: 'aad-3', displayName: 'New User' });
  });

  it('rejects for the reject policy', () => {
    const outcome = resolveIdentity({
      collective: collective([]), state: emptyState(), appId: 'app', aadObjectId: 'aad-4',
      options: options({ unknownUserPolicy: 'reject' }),
    });
    expect(outcome).toEqual({ kind: 'reject', aadObjectId: 'aad-4' });
  });

  it('uses the fallback participant', () => {
    const outcome = resolveIdentity({
      collective: collective([
        { id: 'operator', name: 'Operator', type: 'user', tools: {}, status: 'active' },
      ]),
      state: emptyState(), appId: 'app', aadObjectId: 'aad-5',
      options: options({ unknownUserPolicy: 'fallback', fallbackUserId: 'operator' }),
    });
    expect(outcome).toMatchObject({ kind: 'fallback', user: { participantId: 'operator' } });
  });

  it('treats a retired participant as unknown', () => {
    const outcome = resolveIdentity({
      collective: collective([
        { id: 'old', name: 'Old', type: 'user', tools: {}, status: 'retired',
          identities: [{ connector: 'teams', externalId: 'aad-6' }] },
      ]),
      state: emptyState(), appId: 'app', aadObjectId: 'aad-6',
      options: options({ unknownUserPolicy: 'reject' }),
    });
    expect(outcome.kind).toBe('reject');
  });
});
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `npm test -- src/pairing.test.ts src/identity.test.ts`
Expected: FAIL — modules missing.

- [ ] **Step 3: Implement**

```ts
// src/pairing.ts
import type { PendingPair, TeamsState } from './state.js';

const ALPHABET = '23456789ABCDEFGHJKLMNPQRSTUVWXYZ';

export function generatePairCode(random: () => number = Math.random): string {
  let code = '';
  for (let i = 0; i < 6; i++) code += ALPHABET[Math.floor(random() * ALPHABET.length)];
  return code;
}

export function pruneExpiredPairs(state: TeamsState, now = Date.now()): void {
  for (const [code, pair] of Object.entries(state.pendingPairs)) {
    if (Date.parse(pair.expiresAt) <= now) delete state.pendingPairs[code];
  }
}

export function rememberPendingPair(
  state: TeamsState,
  appId: string,
  aadObjectId: string,
  ttlSeconds: number,
  now = Date.now(),
): string {
  pruneExpiredPairs(state, now);
  for (const [code, pair] of Object.entries(state.pendingPairs)) {
    if (pair.appId === appId && pair.aadObjectId === aadObjectId) return code;
  }
  let code = generatePairCode();
  while (state.pendingPairs[code]) code = generatePairCode();
  state.pendingPairs[code] = {
    appId,
    aadObjectId,
    expiresAt: new Date(now + ttlSeconds * 1000).toISOString(),
  };
  return code;
}

export function consumePendingPair(
  state: TeamsState,
  code: string,
  now = Date.now(),
): PendingPair | null {
  const key = code.toUpperCase();
  const pair = state.pendingPairs[key];
  if (!pair) return null;
  delete state.pendingPairs[key];
  return Date.parse(pair.expiresAt) <= now ? null : pair;
}
```

```ts
// src/identity.ts
import type { CollectiveLike } from '@legion-collective/core';
import type { ResolvedOptions } from './options.js';
import { rememberPendingPair } from './pairing.js';
import type { TeamsState } from './state.js';

export interface ResolvedUser {
  participantId: string;
  aadObjectId: string;
  displayName?: string;
}

export type IdentityOutcome =
  | { kind: 'resolved'; user: ResolvedUser }
  | { kind: 'pairing'; aadObjectId: string; code: string }
  | { kind: 'reject'; aadObjectId: string }
  | { kind: 'auto-provision'; aadObjectId: string; displayName?: string }
  | { kind: 'fallback'; user: ResolvedUser };

export function resolveIdentity(args: {
  collective: CollectiveLike;
  state: TeamsState;
  appId: string;
  aadObjectId: string;
  displayName?: string;
  options: ResolvedOptions;
  now?: number;
}): IdentityOutcome {
  const { collective, state, appId, aadObjectId, displayName, options, now } = args;

  const found = collective.findByIdentity('teams', aadObjectId);
  if (found && (found.status ?? 'active') !== 'retired') {
    return {
      kind: 'resolved',
      user: { participantId: found.id, aadObjectId, displayName },
    };
  }

  if (options.unknownUserPolicy === 'fallback' && options.fallbackUserId) {
    const fallback = collective.get(options.fallbackUserId);
    if (fallback && (fallback.status ?? 'active') !== 'retired') {
      return {
        kind: 'fallback',
        user: { participantId: fallback.id, aadObjectId, displayName },
      };
    }
    return { kind: 'reject', aadObjectId };
  }

  switch (options.unknownUserPolicy) {
    case 'pairing':
      return {
        kind: 'pairing',
        aadObjectId,
        code: rememberPendingPair(state, appId, aadObjectId, options.pairingTtlSeconds, now),
      };
    case 'auto-provision':
      return { kind: 'auto-provision', aadObjectId, displayName };
    case 'reject':
    case 'fallback':
      return { kind: 'reject', aadObjectId };
  }
}
```

- [ ] **Step 4: Run tests**

Run: `npm test -- src/pairing.test.ts src/identity.test.ts`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/pairing.ts src/identity.ts src/pairing.test.ts src/identity.test.ts
git commit -m "feat: add pairing codes and identity resolution"
```

---

## Task 6: Commands

**Files:**
- Create: `src/commands.ts`
- Create: `src/commands.test.ts`

**Interfaces:**
- Produces: `CommandDeps`, `handleCommand(deps): Promise<boolean>`.
- Consumes: `ConnectorContext.callTool`, `consumePendingPair`.

- [ ] **Step 1: Write the failing test**

```ts
// src/commands.test.ts
import { describe, expect, it, vi } from 'vitest';
import { handleCommand } from './commands.js';
import { emptyState } from './state.js';
import { resolveOptions } from './options.js';
import { rememberPendingPair } from './pairing.js';
import type { ConnectorContext } from '@legion-collective/core';

function deps(text: string, overrides: Record<string, unknown> = {}) {
  const sent: string[] = [];
  const callTool = vi.fn(async () => ({ result: { status: 'success' }, conversationId: 'c' }));
  return {
    sent,
    callTool,
    deps: {
      appId: 'app',
      participantId: 'alice',
      aadObjectId: 'aad-1',
      text,
      state: emptyState(),
      ctx: { callTool } as unknown as ConnectorContext,
      options: resolveOptions({ stateDir: '/tmp/x' }),
      send: async (value: string) => { sent.push(value); },
      ...overrides,
    },
  };
}

describe('commands', () => {
  it('clears the Legion conversation on /new', async () => {
    const { deps: d, sent } = deps('/new');
    d.state.legionConversations['app:alice'] = 'legacy';
    await expect(handleCommand(d)).resolves.toBe(true);
    expect(d.state.legionConversations['app:alice']).toBeUndefined();
    expect(sent[0]).toMatch(/new conversation/i);
  });

  it('pairs with a valid code via the core pair tool', async () => {
    const { deps: d, callTool, sent } = deps('/pair ABCDEF bob');
    rememberPendingPair(d.state, 'app', 'aad-1', 600);
    const code = Object.keys(d.state.pendingPairs)[0];
    await handleCommand({ ...d, text: `/pair ${code} bob` });
    expect(callTool).toHaveBeenCalledWith('alice', 'pair', {
      participantId: 'bob',
      connector: 'teams',
      externalId: 'aad-1',
    });
    expect(sent[0]).toMatch(/Paired bob/);
  });

  it('uses pairingParticipantId when configured', async () => {
    const { deps: d, callTool } = deps('/pair', {
      options: resolveOptions({ stateDir: '/tmp/x', pairingParticipantId: 'pairing-agent' }),
    });
    rememberPendingPair(d.state, 'app', 'aad-1', 600);
    const code = Object.keys(d.state.pendingPairs)[0];
    await handleCommand({ ...d, text: `/pair ${code} bob` });
    expect(callTool).toHaveBeenCalledWith('pairing-agent', 'pair', expect.anything());
  });

  it('reports an invalid code', async () => {
    const { deps: d, sent, callTool } = deps('/pair NOPE bob');
    await handleCommand(d);
    expect(callTool).not.toHaveBeenCalled();
    expect(sent[0]).toMatch(/invalid or expired/);
  });

  it('returns false for plain text', async () => {
    await expect(handleCommand(deps('hello').deps)).resolves.toBe(false);
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npm test -- src/commands.test.ts`
Expected: FAIL — module missing.

- [ ] **Step 3: Implement**

```ts
// src/commands.ts
import type { ConnectorContext } from '@legion-collective/core';
import { consumePendingPair } from './pairing.js';
import type { ResolvedOptions } from './options.js';
import type { TeamsState } from './state.js';

export interface CommandDeps {
  appId: string;
  participantId: string;
  aadObjectId: string;
  text: string;
  state: TeamsState;
  ctx: ConnectorContext;
  options: ResolvedOptions;
  send: (text: string) => Promise<void>;
}

export async function handleCommand(deps: CommandDeps): Promise<boolean> {
  const [raw, ...rest] = deps.text.trim().split(/\s+/);
  const command = raw.toLowerCase();

  switch (command) {
    case '/new': {
      delete deps.state.legionConversations[`${deps.appId}:${deps.participantId}`];
      await deps.send('Started a new conversation.');
      return true;
    }
    case '/pair': {
      const code = rest[0]?.toUpperCase();
      const target = rest[1];
      if (!code || !target) {
        await deps.send('Usage: /pair <code> <participant-id>');
        return true;
      }
      const pending = consumePendingPair(deps.state, code);
      if (!pending || pending.appId !== deps.appId || pending.aadObjectId !== deps.aadObjectId) {
        await deps.send('Pairing code is invalid or expired.');
        return true;
      }
      const caller = deps.options.pairingParticipantId ?? deps.participantId;
      const { result } = await deps.ctx.callTool(caller, 'pair', {
        participantId: target,
        connector: 'teams',
        externalId: deps.aadObjectId,
      });
      if (result.status === 'success') {
        await deps.send(`Paired ${target} with your Teams identity.`);
      } else if (result.status === 'pending_approval') {
        await deps.send('Pairing is awaiting approval.');
      } else {
        await deps.send(`Pairing failed: ${result.error ?? result.status}`);
      }
      return true;
    }
    case '/whoami': {
      const conversation =
        deps.state.legionConversations[`${deps.appId}:${deps.participantId}`] ?? '(none)';
      await deps.send(
        `Participant: ${deps.participantId}\nTeams ID: ${deps.aadObjectId}\n` +
        `Bot app: ${deps.appId}\nConversation: ${conversation}`,
      );
      return true;
    }
    case '/help': {
      await deps.send(
        'Commands:\n' +
        '/new — start a new conversation\n' +
        '/pair <code> <participant-id> — pair your Teams identity\n' +
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

- [ ] **Step 4: Run tests**

Run: `npm test -- src/commands.test.ts`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/commands.ts src/commands.test.ts
git commit -m "feat: add /new, /pair, /whoami, /help commands"
```

---

## Task 7: Rendering and Adaptive Cards

**Files:**
- Create: `src/render.ts`
- Create: `src/render.test.ts`

**Interfaces:**
- Produces: `ApprovalRequest`, `chunkMarkdown(text, maxChars)`, `approvalCard(approval)`, `toolCard(name, args)`, `formatToolResult(result)`.

- [ ] **Step 1: Write the failing test**

```ts
// src/render.test.ts
import { describe, expect, it } from 'vitest';
import { approvalCard, chunkMarkdown, formatToolResult, toolCard } from './render.js';

describe('render', () => {
  it('returns a single chunk for short text', () => {
    expect(chunkMarkdown('hello', 100)).toEqual(['hello']);
  });

  it('splits long text on a safe boundary and preserves content', () => {
    const text = `${'a'.repeat(40)}\n${'b'.repeat(40)}\n${'c'.repeat(40)}`;
    const chunks = chunkMarkdown(text, 50);
    expect(chunks.length).toBeGreaterThan(1);
    expect(chunks.join('').replace(/\n/g, '')).toBe(text.replace(/\n/g, ''));
  });

  it('builds an approval card with approve and reject actions', () => {
    const card = approvalCard({
      approvalId: 'ap-1',
      requesterId: 'assistant',
      tool: 'communicate',
      args: { to: 'alice' },
    });
    expect(card.type).toBe('AdaptiveCard');
    expect(card.actions).toHaveLength(2);
    expect(card.actions?.[0]).toMatchObject({
      type: 'Action.Execute',
      title: 'Approve',
      data: { action: 'approve', approvalId: 'ap-1' },
    });
  });

  it('builds a tool card', () => {
    const card = toolCard('communicate', '{"to":"alice"}');
    expect(card.type).toBe('AdaptiveCard');
    expect(JSON.stringify(card)).toContain('communicate');
  });

  it('formats error and rejected results', () => {
    expect(formatToolResult({ status: 'error', error: 'boom' })).toMatch(/boom/);
    expect(formatToolResult({ status: 'rejected', message: 'no' })).toMatch(/no/);
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npm test -- src/render.test.ts`
Expected: FAIL — module missing.

- [ ] **Step 3: Implement**

```ts
// src/render.ts
import type { IAdaptiveCard } from '@microsoft/teams.cards';

export interface ApprovalRequest {
  approvalId: string;
  conversationId?: string;
  requesterId: string;
  tool: string;
  args: unknown;
  createdAt?: string;
}

function truncate(value: string, max: number): string {
  return value.length <= max ? value : `${value.slice(0, max)}…`;
}

export function chunkMarkdown(text: string, maxChars: number): string[] {
  if (text.length <= maxChars) return [text];
  const chunks: string[] = [];
  let remaining = text;
  while (remaining.length > maxChars) {
    let cut = remaining.lastIndexOf('\n', maxChars);
    if (cut < maxChars / 2) cut = remaining.lastIndexOf(' ', maxChars);
    if (cut < maxChars / 2) cut = maxChars;
    chunks.push(remaining.slice(0, cut));
    remaining = remaining.slice(cut).replace(/^\n/, '');
  }
  if (remaining) chunks.push(remaining);
  return chunks;
}

export function approvalCard(approval: ApprovalRequest): IAdaptiveCard {
  return {
    type: 'AdaptiveCard',
    version: '1.4',
    body: [
      { type: 'TextBlock', text: '🔐 Approval requested', weight: 'Bolder', wrap: true },
      {
        type: 'FactSet',
        facts: [
          { title: 'Tool', value: approval.tool },
          { title: 'Args', value: truncate(JSON.stringify(approval.args), 500) },
          { title: 'Requester', value: approval.requesterId },
        ],
      },
    ],
    actions: [
      {
        type: 'Action.Execute',
        title: 'Approve',
        data: { action: 'approve', approvalId: approval.approvalId },
      },
      {
        type: 'Action.Execute',
        title: 'Reject',
        data: { action: 'reject', approvalId: approval.approvalId },
      },
    ],
  } as IAdaptiveCard;
}

export function toolCard(name: string, args: string): IAdaptiveCard {
  return {
    type: 'AdaptiveCard',
    version: '1.4',
    body: [
      { type: 'TextBlock', text: `🛠️ ${name}`, weight: 'Bolder', wrap: true },
      { type: 'TextBlock', text: args || '(no args)', fontType: 'Monospace', wrap: true },
    ],
  } as IAdaptiveCard;
}

export function formatToolResult(result: {
  status: string;
  error?: string;
  message?: string;
  data?: unknown;
}): string {
  if (result.status === 'error') return `Error: ${result.error ?? 'unknown error'}`;
  if (result.status === 'rejected') return result.message ?? 'Rejected.';
  if (result.status === 'pending_approval') return 'Awaiting approval.';
  return '';
}
```

- [ ] **Step 4: Run tests**

Run: `npm test -- src/render.test.ts`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/render.ts src/render.test.ts
git commit -m "feat: render Teams markdown, tool cards, and approval cards"
```

---

## Task 8: `communicate` streaming

**Files:**
- Create: `src/communicate.ts`
- Create: `src/communicate.test.ts`

**Interfaces:**
- Produces: `TeamsStreamer`, `CommunicateDeps`, `runCommunicate(deps): Promise<{ result, conversationId }>`.
- Consumes: `ConnectorContext.streamTool`, `render.approvalCard`, `render.toolCard`.

- [ ] **Step 1: Write the failing test**

```ts
// src/communicate.test.ts
import { describe, expect, it, vi } from 'vitest';
import { runCommunicate, type TeamsStreamer } from './communicate.js';
import { resolveOptions } from './options.js';
import type { ConnectorContext, StreamChunk } from '@legion-collective/core';

function fakeCtx(chunks: StreamChunk[]) {
  return {
    streamTool: vi.fn(async () => ({
      conversationId: 'conv-1',
      gen: (async function* () {
        for (const chunk of chunks) yield chunk;
      })(),
    })),
    callTool: vi.fn(),
  } as unknown as ConnectorContext;
}

function fakeStream(): TeamsStreamer & { emitted: string[]; updates: string[]; closed: number } {
  return {
    emitted: [],
    updates: [],
    closed: 0,
    emit(activity) { this.emitted.push(String(activity)); },
    update(text) { this.updates.push(text); },
    async close() { this.closed++; return { id: 'm1' }; },
  };
}

const options = (overrides = {}) => resolveOptions({ stateDir: '/tmp/x', ...overrides });

describe('runCommunicate', () => {
  it('streams text deltas and closes the stream', async () => {
    const stream = fakeStream();
    const send = vi.fn(async () => undefined);
    const result = await runCommunicate({
      ctx: fakeCtx([
        { type: 'iteration_start', iteration: 1 },
        { type: 'text_delta', delta: 'Hello ' },
        { type: 'text_delta', delta: 'world' },
        { type: 'stream:done', result: { status: 'success', data: { conversationId: 'conv-1' } } },
      ]),
      userId: 'alice', botOwnerId: 'assistant', text: 'hi', options: options(), stream, send,
    });
    expect(result.conversationId).toBe('conv-1');
    expect(stream.emitted).toEqual(['Hello ', 'world']);
    expect(stream.updates).toContain('Thinking…');
    expect(stream.closed).toBe(1);
  });

  it('buffers and sends once for minimal detail', async () => {
    const stream = fakeStream();
    const send = vi.fn(async () => undefined);
    await runCommunicate({
      ctx: fakeCtx([
        { type: 'text_delta', delta: 'Hello ' },
        { type: 'text_delta', delta: 'world' },
        { type: 'stream:done', result: { status: 'success', data: {} } },
      ]),
      userId: 'alice', botOwnerId: 'assistant', text: 'hi',
      options: options({ maxDetail: 'minimal' }), stream, send,
    });
    expect(stream.emitted).toEqual([]);
    expect(send).toHaveBeenCalledTimes(1);
  });

  it('renders approval cards for pending_approval', async () => {
    const send = vi.fn(async () => undefined);
    await runCommunicate({
      ctx: fakeCtx([
        {
          type: 'stream:done',
          result: {
            status: 'pending_approval',
            data: {
              approvalRequests: [
                { approvalId: 'ap-1', requesterId: 'assistant', tool: 'communicate', args: {} },
              ],
            },
          },
        },
      ]),
      userId: 'alice', botOwnerId: 'assistant', text: 'hi',
      options: options(), stream: fakeStream(), send,
    });
    expect(send).toHaveBeenCalledTimes(1);
  });

  it('shows tool status before text', async () => {
    const stream = fakeStream();
    await runCommunicate({
      ctx: fakeCtx([
        { type: 'tool_call_start', index: 0, id: 't1', name: 'search' },
        { type: 'stream:done', result: { status: 'success', data: {} } },
      ]),
      userId: 'alice', botOwnerId: 'assistant', text: 'hi',
      options: options({ toolDisplay: 'status' }), stream, send: async () => undefined,
    });
    expect(stream.updates).toContain('Calling search…');
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npm test -- src/communicate.test.ts`
Expected: FAIL — module missing.

- [ ] **Step 3: Implement**

```ts
// src/communicate.ts
import type { ConnectorContext, StreamChunk, ToolResult } from '@legion-collective/core';
import { MessageActivityInput } from '@microsoft/teams.api';
import type { ResolvedOptions } from './options.js';
import { approvalCard, toolCard, type ApprovalRequest } from './render.js';

export interface TeamsStreamer {
  emit(activity: MessageActivityInput | string): void;
  update(text: string): void;
  close(): Promise<{ id?: string } | undefined>;
}

export interface CommunicateDeps {
  ctx: ConnectorContext;
  userId: string;
  botOwnerId: string;
  text: string;
  conversationId?: string;
  options: ResolvedOptions;
  stream?: TeamsStreamer;
  send: (activity: MessageActivityInput) => Promise<unknown>;
}

export async function runCommunicate(deps: CommunicateDeps): Promise<{
  result: ToolResult;
  conversationId: string;
}> {
  const { gen, conversationId } = await deps.ctx.streamTool(deps.userId, 'communicate', {
    to: deps.botOwnerId,
    message: deps.text,
    detail: deps.options.maxDetail,
    conversationId: deps.conversationId,
  });

  let buffer = '';
  let final: ToolResult | undefined;
  let sawError: string | undefined;

  for await (const chunk of gen as AsyncGenerator<StreamChunk>) {
    switch (chunk.type) {
      case 'iteration_start':
        if (deps.options.maxDetail !== 'minimal') deps.stream?.update('Thinking…');
        break;
      case 'text_delta':
        buffer += chunk.delta;
        if (deps.options.maxDetail !== 'minimal') deps.stream?.emit(chunk.delta);
        break;
      case 'message_snapshot':
        buffer = chunk.content;
        break;
      case 'tool_call_start':
        if (deps.options.toolDisplay === 'status') {
          deps.stream?.update(`Calling ${chunk.name}…`);
        } else if (deps.options.toolDisplay === 'cards') {
          await deps.send(
            new MessageActivityInput().addCard('adaptive', toolCard(chunk.name, '')),
          );
        }
        break;
      case 'call_timeout':
        await deps.send(new MessageActivityInput(`Timed out after ${chunk.timeoutMs} ms.`));
        break;
      case 'stream:done':
        final = chunk.result;
        break;
      case 'stream:error':
        sawError = chunk.error;
        break;
      default:
        break;
    }
  }

  if (sawError) {
    await deps.send(new MessageActivityInput(`Error: ${sawError}`));
    final ??= { status: 'error', error: sawError };
  }

  if (deps.options.maxDetail === 'minimal') {
    if (buffer) await deps.send(new MessageActivityInput(buffer));
  } else {
    await deps.stream?.close();
  }

  if (final?.status === 'pending_approval') {
    const data = final.data as { approvalRequests?: ApprovalRequest[] } | undefined;
    for (const approval of data?.approvalRequests ?? []) {
      await deps.send(new MessageActivityInput().addCard('adaptive', approvalCard(approval)));
    }
  } else if (final?.status === 'error') {
    await deps.send(new MessageActivityInput(`Error: ${final.error ?? 'unknown error'}`));
  } else if (final?.status === 'rejected') {
    await deps.send(new MessageActivityInput(final.message ?? 'Rejected.'));
  }

  return { result: final ?? { status: 'success' }, conversationId };
}
```

- [ ] **Step 4: Run tests**

Run: `npm test -- src/communicate.test.ts`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/communicate.ts src/communicate.test.ts
git commit -m "feat: stream communicate output into Teams"
```

---

## Task 9: Approvals

**Files:**
- Create: `src/approvals.ts`
- Create: `src/approvals.test.ts`

**Interfaces:**
- Produces: `sendApprovalCards(approvals, send)`, `handleApprovalAction({ ctx, userId, decision, approvalId })`.
- Consumes: `ConnectorContext.callTool`, `render.approvalCard`.

- [ ] **Step 1: Write the failing test**

```ts
// src/approvals.test.ts
import { describe, expect, it, vi } from 'vitest';
import { handleApprovalAction, sendApprovalCards } from './approvals.js';
import type { ConnectorContext } from '@legion-collective/core';

describe('approvals', () => {
  it('sends one card per approval request', async () => {
    const send = vi.fn(async () => undefined);
    await sendApprovalCards(
      [
        { approvalId: 'ap-1', requesterId: 'assistant', tool: 'communicate', args: {} },
        { approvalId: 'ap-2', requesterId: 'researcher', tool: 'communicate', args: {} },
      ],
      send,
    );
    expect(send).toHaveBeenCalledTimes(2);
  });

  it('calls approval_response as the clicking user', async () => {
    const callTool = vi.fn(async () => ({ result: { status: 'success' }, conversationId: 'c' }));
    const ctx = { callTool } as unknown as ConnectorContext;
    const result = await handleApprovalAction({
      ctx, userId: 'alice', decision: 'approve', approvalId: 'ap-1',
    });
    expect(callTool).toHaveBeenCalledWith('alice', 'approval_response', {
      decisions: [{ approvalId: 'ap-1', decision: 'approve' }],
    });
    expect(result.status).toBe('success');
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npm test -- src/approvals.test.ts`
Expected: FAIL — module missing.

- [ ] **Step 3: Implement**

```ts
// src/approvals.ts
import type { ConnectorContext, ToolResult } from '@legion-collective/core';
import { MessageActivityInput } from '@microsoft/teams.api';
import { approvalCard, type ApprovalRequest } from './render.js';

export async function sendApprovalCards(
  approvals: ApprovalRequest[],
  send: (activity: MessageActivityInput) => Promise<unknown>,
): Promise<void> {
  for (const approval of approvals) {
    await send(new MessageActivityInput().addCard('adaptive', approvalCard(approval)));
  }
}

export async function handleApprovalAction(args: {
  ctx: ConnectorContext;
  userId: string;
  decision: 'approve' | 'reject';
  approvalId: string;
}): Promise<ToolResult> {
  const { result } = await args.ctx.callTool(args.userId, 'approval_response', {
    decisions: [{ approvalId: args.approvalId, decision: args.decision }],
  });
  return result;
}
```

- [ ] **Step 4: Run tests**

Run: `npm test -- src/approvals.test.ts`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/approvals.ts src/approvals.test.ts
git commit -m "feat: render and handle Teams approval cards"
```

---

## Task 10: Inbound handler

**Files:**
- Create: `src/inbound.ts`
- Create: `src/inbound.test.ts`

**Interfaces:**
- Produces: `TeamsActivity`, `TeamsContextLike`, `InboundDeps`, `handleInbound(deps, context)`.
- Consumes: `resolveIdentity`, `handleCommand`, `runCommunicate`, `StateStore`, `ConnectorContext`.

- [ ] **Step 1: Write the failing test**

```ts
// src/inbound.test.ts
import { describe, expect, it, vi } from 'vitest';
import { handleInbound } from './inbound.js';
import { emptyState } from './state.js';
import { resolveOptions } from './options.js';
import type { BotEntry } from './directory.js';
import type { ConnectorContext, CollectiveLike } from '@legion-collective/core';

function setup(participants: Array<Record<string, unknown>> = []) {
  const state = emptyState();
  const save = vi.fn(async () => undefined);
  const callTool = vi.fn(async () => ({ result: { status: 'success' }, conversationId: 'c' }));
  const streamTool = vi.fn(async () => ({
    conversationId: 'conv-1',
    gen: (async function* () {
      yield { type: 'text_delta', delta: 'Hi' };
      yield { type: 'stream:done', result: { status: 'success', data: {} } };
    })(),
  }));
  const collective = {
    list: () => participants,
    get: (id: string) => participants.find((p) => p.id === id),
    findByIdentity: (connector: string, externalId: string) =>
      participants.find((p) =>
        (p.identities as Array<{ connector: string; externalId: string }> | undefined)?.some(
          (i) => i.connector === connector && i.externalId === externalId,
        ),
      ),
  } as unknown as CollectiveLike;
  const bot: BotEntry = {
    participantId: 'assistant', name: 'Assistant', appId: 'app',
    credentials: { appId: 'app', clientSecret: 's', tenantId: 't' },
  };
  return {
    state, save, callTool, streamTool, collective, bot,
    ctx: { callTool, streamTool, registry: { setActive: vi.fn() } } as unknown as ConnectorContext,
  };
}

const activity = (text: string, aad = 'aad-1') => ({
  from: { id: '29:1', name: 'Alice', aadObjectId: aad },
  conversation: { id: 'conv-1' },
  channelData: { tenant: { id: 'tenant-1' } },
  serviceUrl: 'https://smba.trafficmanager.net/teams/',
  text,
});

describe('handleInbound', () => {
  it('streams communicate for a resolved user and stores state', async () => {
    const s = setup([
      { id: 'alice', name: 'Alice', type: 'user', tools: {}, status: 'active',
        identities: [{ connector: 'teams', externalId: 'aad-1' }] },
    ]);
    const send = vi.fn(async () => undefined);
    await handleInbound(
      { bot: s.bot, ctx: s.ctx, collective: s.collective, state: s.state,
        stateStore: { save: s.save } as never, options: resolveOptions({ stateDir: '/tmp/x' }) },
      { activity: activity('hello'), send },
    );
    expect(s.streamTool).toHaveBeenCalledWith('alice', 'communicate', expect.objectContaining({
      to: 'assistant', message: 'hello',
    }));
    expect(s.state.legionConversations['app:alice']).toBe('conv-1');
    expect(s.state.conversations['app:alice'].conversationId).toBe('conv-1');
    expect(s.save).toHaveBeenCalled();
  });

  it('issues a pairing code for an unknown user', async () => {
    const s = setup();
    const send = vi.fn(async () => undefined);
    await handleInbound(
      { bot: s.bot, ctx: s.ctx, collective: s.collective, state: s.state,
        stateStore: { save: s.save } as never, options: resolveOptions({ stateDir: '/tmp/x' }) },
      { activity: activity('hello', 'aad-9'), send },
    );
    expect(send.mock.calls[0][0]).toMatch(/Not registered/);
    expect(Object.keys(s.state.pendingPairs)).toHaveLength(1);
  });

  it('handles /new without calling communicate', async () => {
    const s = setup([
      { id: 'alice', name: 'Alice', type: 'user', tools: {}, status: 'active',
        identities: [{ connector: 'teams', externalId: 'aad-1' }] },
    ]);
    s.state.legionConversations['app:alice'] = 'old';
    const send = vi.fn(async () => undefined);
    await handleInbound(
      { bot: s.bot, ctx: s.ctx, collective: s.collective, state: s.state,
        stateStore: { save: s.save } as never, options: resolveOptions({ stateDir: '/tmp/x' }) },
      { activity: activity('/new'), send },
    );
    expect(s.state.legionConversations['app:alice']).toBeUndefined();
    expect(s.streamTool).not.toHaveBeenCalled();
  });

  it('auto-provisions then pairs an unknown user', async () => {
    const s = setup();
    const send = vi.fn(async () => undefined);
    await handleInbound(
      { bot: s.bot, ctx: s.ctx, collective: s.collective, state: s.state,
        stateStore: { save: s.save } as never,
        options: resolveOptions({ stateDir: '/tmp/x', unknownUserPolicy: 'auto-provision' }) },
      { activity: activity('hello', 'aad-7'), send },
    );
    expect(s.callTool).toHaveBeenCalledWith('operator', 'create_user', expect.objectContaining({
      name: 'Alice',
    }));
    expect(s.callTool).toHaveBeenCalledWith('operator', 'pair', expect.objectContaining({
      connector: 'teams', externalId: 'aad-7',
    }));
    expect(send.mock.calls[0][0]).toMatch(/Welcome/);
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npm test -- src/inbound.test.ts`
Expected: FAIL — module missing.

- [ ] **Step 3: Implement**

```ts
// src/inbound.ts
import type { CollectiveLike, ConnectorContext } from '@legion-collective/core';
import { MessageActivityInput, TypingActivityInput } from '@microsoft/teams.api';
import { handleCommand } from './commands.js';
import { runCommunicate, type TeamsStreamer } from './communicate.js';
import type { BotEntry } from './directory.js';
import { resolveIdentity } from './identity.js';
import type { ResolvedOptions } from './options.js';
import type { StateStore, TeamsState } from './state.js';

export interface TeamsActivity {
  from?: { id?: string; name?: string; aadObjectId?: string };
  conversation?: { id?: string };
  channelData?: { tenant?: { id?: string } };
  serviceUrl?: string;
  text?: string;
  value?: { action?: { data?: Record<string, unknown> } };
}

export interface TeamsContextLike {
  activity: TeamsActivity;
  send: (activity: MessageActivityInput | TypingActivityInput | string) => Promise<unknown>;
  stream?: TeamsStreamer;
}

export interface InboundDeps {
  bot: BotEntry;
  ctx: ConnectorContext;
  collective: CollectiveLike;
  state: TeamsState;
  stateStore: StateStore;
  options: ResolvedOptions;
}

function slug(value: string): string {
  return value.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '').slice(0, 24) || 'user';
}

function hash6(value: string): string {
  let hash = 0;
  for (const char of value) hash = (hash * 31 + char.charCodeAt(0)) >>> 0;
  return hash.toString(36).padStart(6, '0').slice(0, 6);
}

function upsertConversation(
  state: TeamsState,
  bot: BotEntry,
  userId: string,
  aadObjectId: string,
  displayName: string | undefined,
  activity: TeamsActivity,
): void {
  state.conversations[`${bot.appId}:${userId}`] = {
    conversationId: activity.conversation?.id ?? '',
    tenantId: activity.channelData?.tenant?.id,
    serviceUrl: activity.serviceUrl,
    aadObjectId,
    displayName,
    updatedAt: new Date().toISOString(),
  };
}

async function autoProvision(
  deps: InboundDeps,
  aadObjectId: string,
  displayName: string | undefined,
  context: TeamsContextLike,
): Promise<void> {
  const name = displayName ?? 'Teams user';
  const id = `teams-${slug(name)}-${hash6(aadObjectId)}`;

  const created = await deps.ctx.callTool(deps.options.provisionerId, 'create_user', { id, name });
  if (created.result.status !== 'success') {
    await context.send(`Provisioning failed: ${created.result.error ?? created.result.status}`);
    return;
  }

  const paired = await deps.ctx.callTool(
    deps.options.pairingParticipantId ?? deps.options.provisionerId,
    'pair',
    { participantId: id, connector: 'teams', externalId: aadObjectId },
  );
  if (paired.result.status !== 'success') {
    await context.send(`Pairing failed: ${paired.result.error ?? paired.result.status}`);
    return;
  }

  await context.send(`Welcome, ${name}. You are now paired as ${id}.`);
}

export async function handleInbound(
  deps: InboundDeps,
  context: TeamsContextLike,
): Promise<void> {
  const activity = context.activity;
  const aadObjectId = activity.from?.aadObjectId ?? activity.from?.id;
  if (!aadObjectId) {
    await context.send('Unable to read your Teams identity.');
    return;
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
      await context.send(
        `Not registered with Legion yet. Ask someone with the \`pair\` tool to run ` +
        `/pair ${outcome.code} <participant-id>. Your Teams ID: ${outcome.aadObjectId}.`,
      );
      await deps.stateStore.save(deps.state);
      return;

    case 'reject':
      await context.send(
        `Not registered with Legion yet. Your Teams ID: ${outcome.aadObjectId}.`,
      );
      return;

    case 'auto-provision':
      await autoProvision(deps, outcome.aadObjectId, outcome.displayName, context);
      await deps.stateStore.save(deps.state);
      return;

    case 'resolved':
    case 'fallback': {
      const user = outcome.user;
      upsertConversation(
        deps.state,
        deps.bot,
        user.participantId,
        user.aadObjectId,
        user.displayName,
        activity,
      );
      await deps.stateStore.save(deps.state);
      deps.ctx.registry.setActive(user.participantId, 'teams');

      const text = activity.text?.trim();
      if (!text) {
        await context.send('Empty message.');
        return;
      }

      if (text.startsWith('/')) {
        const handled = await handleCommand({
          appId: deps.bot.appId,
          participantId: user.participantId,
          aadObjectId: user.aadObjectId,
          text,
          state: deps.state,
          ctx: deps.ctx,
          options: deps.options,
          send: (value) => context.send(value),
        });
        if (handled) {
          await deps.stateStore.save(deps.state);
          return;
        }
      }

      let typing: NodeJS.Timeout | undefined;
      if (deps.options.typingIntervalMs > 0) {
        const tick = () => context.send(new TypingActivityInput()).catch(() => undefined);
        tick();
        typing = setInterval(tick, deps.options.typingIntervalMs);
      }

      try {
        const { conversationId } = await runCommunicate({
          ctx: deps.ctx,
          userId: user.participantId,
          botOwnerId: deps.bot.participantId,
          text,
          conversationId:
            deps.state.legionConversations[`${deps.bot.appId}:${user.participantId}`],
          options: deps.options,
          stream: context.stream,
          send: (activity) => context.send(activity),
        });
        deps.state.legionConversations[`${deps.bot.appId}:${user.participantId}`] =
          conversationId;
        await deps.stateStore.save(deps.state);
      } finally {
        if (typing) clearInterval(typing);
      }
      return;
    }
  }
}
```

- [ ] **Step 4: Run tests**

Run: `npm test -- src/inbound.test.ts`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/inbound.ts src/inbound.test.ts
git commit -m "feat: handle inbound Teams messages, commands, and provisioning"
```

---

## Task 11: Outbound delivery

**Files:**
- Create: `src/outbound.ts`
- Create: `src/outbound.test.ts`

**Interfaces:**
- Produces: `OutboundApp`, `OutboundDeps`, `deliverMessage(deps, message)`.
- Consumes: `CollectiveLike.get`, `StateStore.save`, `render.chunkMarkdown`.

- [ ] **Step 1: Write the failing test**

```ts
// src/outbound.test.ts
import { describe, expect, it, vi } from 'vitest';
import { deliverMessage } from './outbound.js';
import { emptyState } from './state.js';
import { resolveOptions } from './options.js';
import type { CollectiveLike } from '@legion-collective/core';

function setup(participants: Array<Record<string, unknown>>, conversations: Record<string, unknown> = {}) {
  const state = { ...emptyState(), conversations };
  const send = vi.fn(async () => undefined);
  const create = vi.fn(async () => ({ id: 'conv-new' }));
  const save = vi.fn(async () => undefined);
  const collective = {
    list: () => participants,
    get: (id: string) => participants.find((p) => p.id === id),
  } as unknown as CollectiveLike;
  const apps = {
    getByAppId: (appId: string) => (appId === 'app' ? { send, api: { conversations: { create } } } : undefined),
  };
  return { state, send, create, save, collective, apps };
}

describe('deliverMessage', () => {
  it('sends through the sender bot to a stored conversation', async () => {
    const s = setup(
      [
        { id: 'assistant', name: 'Assistant', type: 'agent', tools: {},
          identities: [{ connector: 'teams-bot', externalId: 'app' }] },
        { id: 'alice', name: 'Alice', type: 'user', tools: {},
          identities: [{ connector: 'teams', externalId: 'aad-1' }] },
      ],
      { 'app:alice': { conversationId: 'conv-1', updatedAt: 'now' } },
    );
    await deliverMessage(
      { collective: s.collective, state: s.state as never, stateStore: { save: s.save } as never,
        apps: s.apps as never, options: resolveOptions({ stateDir: '/tmp/x', tenantId: 'tenant-1' }) },
      { senderId: 'assistant', recipientId: 'alice', content: 'hello' },
    );
    expect(s.send).toHaveBeenCalledWith('conv-1', expect.anything());
    expect(s.create).not.toHaveBeenCalled();
  });

  it('creates a personal conversation when none is stored', async () => {
    const s = setup([
      { id: 'assistant', name: 'Assistant', type: 'agent', tools: {},
        identities: [{ connector: 'teams-bot', externalId: 'app' }] },
      { id: 'alice', name: 'Alice', type: 'user', tools: {},
        identities: [{ connector: 'teams', externalId: 'aad-1' }] },
    ]);
    await deliverMessage(
      { collective: s.collective, state: s.state as never, stateStore: { save: s.save } as never,
        apps: s.apps as never, options: resolveOptions({ stateDir: '/tmp/x', tenantId: 'tenant-1' }) },
      { senderId: 'assistant', recipientId: 'alice', content: 'hello' },
    );
    expect(s.create).toHaveBeenCalledWith(expect.objectContaining({
      members: [{ id: 'aad-1' }],
      tenantId: 'tenant-1',
    }));
    expect(s.send).toHaveBeenCalledWith('conv-new', expect.anything());
  });

  it('drops when the sender has no Teams bot', async () => {
    const s = setup([
      { id: 'assistant', name: 'Assistant', type: 'agent', tools: {} },
      { id: 'alice', name: 'Alice', type: 'user', tools: {},
        identities: [{ connector: 'teams', externalId: 'aad-1' }] },
    ]);
    await deliverMessage(
      { collective: s.collective, state: s.state as never, stateStore: { save: s.save } as never,
        apps: s.apps as never, options: resolveOptions({ stateDir: '/tmp/x' }) },
      { senderId: 'assistant', recipientId: 'alice', content: 'hello' },
    );
    expect(s.send).not.toHaveBeenCalled();
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npm test -- src/outbound.test.ts`
Expected: FAIL — module missing.

- [ ] **Step 3: Implement**

```ts
// src/outbound.ts
import type { CollectiveLike } from '@legion-collective/core';
import { MessageActivityInput } from '@microsoft/teams.api';
import type { ResolvedOptions } from './options.js';
import { chunkMarkdown } from './render.js';
import type { StateStore, TeamsState } from './state.js';

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
  apps: { getByAppId(appId: string): OutboundApp | undefined };
  options: ResolvedOptions;
}

export async function deliverMessage(
  deps: OutboundDeps,
  message: { senderId: string; recipientId: string; content: string },
): Promise<void> {
  const sender = deps.collective.get(message.senderId);
  const botIdentity = sender?.identities?.find((i) => i.connector === 'teams-bot');
  if (!botIdentity) return;

  const app = deps.apps.getByAppId(botIdentity.externalId);
  if (!app) return;

  const key = `${botIdentity.externalId}:${message.recipientId}`;
  let record = deps.state.conversations[key];

  if (!record) {
    const recipient = deps.collective.get(message.recipientId);
    const userIdentity = recipient?.identities?.find((i) => i.connector === 'teams');
    if (!userIdentity) return;

    const conversation = await app.api.conversations.create({
      members: [{ id: userIdentity.externalId }],
      tenantId: deps.options.tenantId,
      channelData: { tenant: { id: deps.options.tenantId } },
    });
    record = {
      conversationId: conversation.id,
      tenantId: deps.options.tenantId,
      aadObjectId: userIdentity.externalId,
      updatedAt: new Date().toISOString(),
    };
    deps.state.conversations[key] = record;
    await deps.stateStore.save(deps.state);
  }

  for (const chunk of chunkMarkdown(message.content, deps.options.maxMessageChars)) {
    const activity = new MessageActivityInput(chunk);
    if (deps.options.aiLabel) activity.addAiGenerated();
    await app.send(record.conversationId, activity);
  }
}
```

- [ ] **Step 4: Run tests**

Run: `npm test -- src/outbound.test.ts`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/outbound.ts src/outbound.test.ts
git commit -m "feat: deliver Legion messages through Teams bots"
```

---

## Task 12: Bot pool, transports, and connector wiring

**Files:**
- Create: `src/bot-pool.ts`
- Modify: `src/connector.ts` (full implementation)
- Create: `src/bot-pool.test.ts`, `src/connector.test.ts`

**Interfaces:**
- Produces: `TeamsAppLike`, `AppHandlers`, `TeamsAppFactory`, `defaultAppFactory`, `BotPool` (`start`, `reconcile`, `getByAppId`, `stop`).
- Produces: `TeamsConnector` constructor `(options, deps, appFactory?)`.

- [ ] **Step 1: Write the failing tests**

```ts
// src/bot-pool.test.ts
import { describe, expect, it, vi } from 'vitest';
import { BotPool, type TeamsAppLike } from './bot-pool.js';
import { resolveOptions } from './options.js';
import type { BotEntry } from './directory.js';

function fakeApp(): TeamsAppLike & { sent: unknown[]; stopped: number; initialized: number } {
  return {
    sent: [],
    stopped: 0,
    initialized: 0,
    async initialize() { this.initialized++; },
    async start() { this.initialized++; },
    async stop() { this.stopped++; },
    async send(_conversationId: string, activity: unknown) { this.sent.push(activity); },
    api: { conversations: { create: async () => ({ id: 'conv' }) } },
    on() {},
  };
}

const bot = (participantId: string, appId: string): BotEntry => ({
  participantId, name: participantId, appId,
  credentials: { appId, clientSecret: 's', tenantId: 't' },
});

const handlers = {
  onMessage: vi.fn(async () => undefined),
  onInstall: vi.fn(async () => undefined),
  onConversationUpdate: vi.fn(async () => undefined),
  onApproval: vi.fn(async () => undefined),
};

describe('BotPool', () => {
  it('starts an app per bot and stops them', async () => {
    const apps: Array<ReturnType<typeof fakeApp>> = [];
    const pool = new BotPool(
      resolveOptions({ stateDir: '/tmp/x', port: 0 }),
      (entry) => { const app = fakeApp(); apps.push(app); return app; },
      handlers,
    );
    await pool.start(new Map([['app', bot('assistant', 'app')]]));
    expect(apps).toHaveLength(1);
    expect(pool.getByAppId('app')).toBe(apps[0]);
    await pool.stop();
    expect(apps[0].stopped).toBe(1);
  });

  it('reconciles socket apps individually', async () => {
    const apps = new Map<string, ReturnType<typeof fakeApp>>();
    const pool = new BotPool(
      resolveOptions({ stateDir: '/tmp/x', transport: 'socket' }),
      (entry) => { const app = fakeApp(); apps.set(entry.appId, app); return app; },
      handlers,
    );
    await pool.start(new Map([['a', bot('a', 'a')]]));
    await pool.reconcile(new Map([['b', bot('b', 'b')]]));
    expect(pool.getByAppId('a')).toBeUndefined();
    expect(pool.getByAppId('b')).toBeDefined();
  });
});
```

```ts
// src/connector.test.ts
import { mkdtemp } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { describe, expect, it, vi } from 'vitest';
import { TeamsConnector } from './connector.js';
import type { CollectiveLike, ConnectorRuntimeDeps } from '@legion-collective/core';
import type { TeamsAppLike } from './bot-pool.js';

function setup() {
  const sent: Array<{ conversationId: string; activity: unknown }> = [];
  const app: TeamsAppLike = {
    async initialize() {},
    async start() {},
    async stop() {},
    async send(conversationId, activity) { sent.push({ conversationId, activity }); },
    api: { conversations: { create: async () => ({ id: 'conv-new' }) } },
    on() {},
  };
  const participants = [
    { id: 'assistant', name: 'Assistant', type: 'agent', tools: {}, status: 'active',
      identities: [{ connector: 'teams-bot', externalId: 'app' }] },
    { id: 'alice', name: 'Alice', type: 'user', tools: {}, status: 'active',
      identities: [{ connector: 'teams', externalId: 'aad-1' }] },
  ];
  const collective = {
    list: () => participants,
    get: (id: string) => participants.find((p) => p.id === id),
    findByIdentity: (connector: string, externalId: string) =>
      participants.find((p) =>
        p.identities?.some((i) => i.connector === connector && i.externalId === externalId)),
  } as unknown as CollectiveLike;
  const deps: ConnectorRuntimeDeps = {
    collective,
    secrets: {
      getSecret: async (id: string, key: string) =>
        id === 'assistant' && key === 'teams-bot/app'
          ? '{"appId":"app","clientSecret":"s","tenantId":"t"}'
          : null,
    },
  };
  return { app, sent, deps };
}

describe('TeamsConnector', () => {
  it('starts the pool and delivers outbound messages', async () => {
    const s = setup();
    const dir = await mkdtemp(join(tmpdir(), 'teams-conn-'));
    const c = new TeamsConnector(
      { stateDir: dir, port: 0, tenantId: 'tenant-1' },
      s.deps,
      () => s.app,
    );
    await c.start({ registry: { setActive() {}, clearActive() {} } } as never);
    await c.deliver({
      id: 'm1', conversationId: 'conv-1', senderId: 'assistant',
      recipientId: 'alice', content: 'hello', timestamp: 'now',
    });
    expect(s.sent).toHaveLength(1);
    expect(s.sent[0].conversationId).toBe('conv-new');
    await c.stop();
  });
});
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `npm test -- src/bot-pool.test.ts src/connector.test.ts`
Expected: FAIL — modules incomplete.

- [ ] **Step 3: Implement `bot-pool.ts`**

```ts
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
  app.on('card.action.approve', (ctx) => handlers.onApproval(entry, ctx as never, 'approve'));
  app.on('card.action.reject', (ctx) => handlers.onApproval(entry, ctx as never, 'reject'));

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
```

- [ ] **Step 4: Implement the full `connector.ts`**

```ts
// src/connector.ts
import type { Connector, ConnectorContext, ConnectorRuntimeDeps } from '@legion-collective/core';
import { handleApprovalAction } from './approvals.js';
import { BotPool, defaultAppFactory, type TeamsAppFactory } from './bot-pool.js';
import { buildDirectory, type BotEntry } from './directory.js';
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
    if (!this.deps.secrets) return;
    const directory = await buildDirectory(this.deps.collective, this.deps.secrets);
    for (const warning of directory.warnings) console.warn(`[teams] ${warning}`);
    await this.pool?.reconcile(directory.bots);
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
    const aadObjectId = activity.from?.aadObjectId ?? activity.from?.id;
    if (!aadObjectId) return;
    const participant = this.deps.collective.findByIdentity('teams', aadObjectId);
    if (!participant) return;
    this.state.conversations[`${bot.appId}:${participant.id}`] = {
      conversationId: activity.conversation?.id ?? '',
      tenantId: activity.channelData?.tenant?.id,
      serviceUrl: activity.serviceUrl,
      aadObjectId,
      displayName: participant.name,
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
    if (this.rescan) clearInterval(this.rescan);
    await this.pool?.stop();
    this.ctx = undefined;
  }
}
```

- [ ] **Step 5: Run tests**

Run: `npm test -- src/bot-pool.test.ts src/connector.test.ts`
Expected: PASS.

- [ ] **Step 6: Commit**

```bash
git add src/bot-pool.ts src/connector.ts src/bot-pool.test.ts src/connector.test.ts
git commit -m "feat: run Teams bots over HTTP or Socket Mode"
```

---

## Task 13: Provisioning script

**Files:**
- Create: `scripts/provision-teams-bot.mjs`
- Create: `scripts/provision-all.mjs`

**Interfaces:**
- CLI: `--participant <id> --base-url <url> [--env .env.teams] [--workspace <path>] [--legion-url <url>] [--legion-password <pw>] [--dry-run]`.
- Writes the manifest and zip, stores the SecretStore value, and prints the install link.

- [ ] **Step 1: Write the script**

```js
#!/usr/bin/env node
// scripts/provision-teams-bot.mjs
import { execFileSync } from 'node:child_process';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { join, resolve } from 'node:path';

const COLOR_PNG =
  'iVBORw0KGgoAAAANSUhEUgAAAMAAAADACAYAAABS3GwHAAABiklEQVR42u3TMQ0AAAjAMGRiAP8JBsAEHz1qYMkiqwe+ChEwABgADAAGAAOAAcAAYAAwABgADAAGAAOAAcAAYAAwABgADAAGAAOAAcAAYAAwABgADAAGAAOAAcAAYAAwABgADAAGAAOAAcAAYAAwABgADAAGAAOAAcAAYAAwABgADIABRMAAYAAwABgADAAGAAOAAcAAYAAwABgADAAGAAOAAcAAYAAwABgADAAGAAOAAcAAYAAwABgADAAGAAOAAcAAYAAwABgADAAGAAOAAcAAYAAwABgADAAGAAOAAcAAYAAwAAYAA4ABwABgADAAGAAMAAYAA4ABwABgADAAGAAMAAYAA4ABwABgADAAGAAMAAYAA4ABwABgADAAGAAMAAYAA4ABwABgADAAGAAMAAYAA4ABwABgADAAGAAMAAYAA4ABMIAIGAAMAAYAA4ABwABgADAAGAAMAAYAA4ABwABgADAAGAAMAAYAA4ABwABgADAAGAAMAAYAA4ABwABgADAAGAAMAAaASwtC9faQjfST9wAAAABJRU5ErkJggg==';
const OUTLINE_PNG =
  'iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAAKUlEQVR42u3OIQEAAAACIP+f1hkWWEB6FgEBAQEBAQEBAQEBAQEBgXdgl/rw4tnPBf0AAAAASUVORK5CYII=';

function arg(name, fallback) {
  const index = process.argv.indexOf(`--${name}`);
  return index >= 0 ? process.argv[index + 1] : fallback;
}

const participantId = arg('participant');
const baseUrl = arg('base-url');
if (!participantId || !baseUrl) {
  console.error('Usage: provision-teams-bot.mjs --participant <id> --base-url <url>');
  process.exit(2);
}

const workspace = resolve(arg('workspace', process.env.LEGION_WORKSPACE ?? process.cwd()));
const envFile = resolve(arg('env', join(workspace, '.env.teams')));
const legionUrl = arg('legion-url', 'http://127.0.0.1:3000');
const legionPassword = arg('legion-password', process.env.LEGION_BOOTSTRAP_PASSWORD);
const dryRun = process.argv.includes('--dry-run');

const participantPath = join(workspace, '.legion/collective/participants', `${participantId}.json`);
const participant = JSON.parse(await readFile(participantPath, 'utf8'));
const endpoint = `${baseUrl.replace(/\/$/, '')}/teams/${participantId}/api/messages`;

if (!dryRun) {
  execFileSync('teams', [
    'app', 'create',
    '--name', participant.name,
    '--endpoint', endpoint,
    '--env', envFile,
  ], { stdio: 'inherit' });
}

const env = {};
if (!dryRun) {
  for (const line of (await readFile(envFile, 'utf8')).split(/\r?\n/)) {
    const match = line.match(/^([A-Z0-9_]+)=(.*)$/);
    if (match) env[match[1]] = match[2];
  }
}
const appId = env.CLIENT_ID ?? '<app-id>';
const clientSecret = env.CLIENT_SECRET ?? '<client-secret>';
const tenantId = env.TENANT_ID ?? '<tenant-id>';

const manifest = {
  $schema: 'https://developer.microsoft.com/en-us/json-schemas/teams/v1.19/MicrosoftTeams.schema.json',
  manifestVersion: '1.19',
  version: '1.0.0',
  id: appId,
  packageName: `com.prograde-labs.legion.${participantId}`,
  developer: {
    name: 'Prograde Labs',
    websiteUrl: 'https://prograde-labs.com',
    privacyUrl: 'https://prograde-labs.com',
    termsOfUseUrl: 'https://prograde-labs.com',
  },
  icons: { color: 'color.png', outline: 'outline.png' },
  name: { short: participant.name, full: participant.name },
  description: {
    short: `Legion agent ${participant.name}`,
    full: `Legion agent ${participant.name} on Microsoft Teams.`,
  },
  accentColor: '#5865F2',
  bots: [{ botId: appId, scopes: ['personal'], supportsFiles: false, isNotificationOnly: false }],
  permissions: ['identity', 'messageTeamMembers'],
  validDomains: [],
};

const outDir = join(workspace, '.legion/teams-apps', participantId);
await mkdir(outDir, { recursive: true });
await writeFile(join(outDir, 'manifest.json'), JSON.stringify(manifest, null, 2));
await writeFile(join(outDir, 'color.png'), Buffer.from(COLOR_PNG, 'base64'));
await writeFile(join(outDir, 'outline.png'), Buffer.from(OUTLINE_PNG, 'base64'));

if (!dryRun) {
  execFileSync('zip', ['-j', join(outDir, `${participantId}.zip`), 'manifest.json', 'color.png', 'outline.png'], { cwd: outDir, stdio: 'inherit' });
}

const botIdentity = { connector: 'teams-bot', externalId: appId };
const identities = participant.identities ?? [];
if (!identities.some((i) => i.connector === botIdentity.connector && i.externalId === botIdentity.externalId)) {
  participant.identities = [...identities, botIdentity];
  await writeFile(participantPath, JSON.stringify(participant, null, 2));
}

if (!dryRun) {
  const login = await fetch(`${legionUrl}/api/auth/login`, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ name: 'operator', password: legionPassword }),
  });
  const { token } = await login.json();
  const response = await fetch(`${legionUrl}/api/execute`, {
    method: 'POST',
    headers: { authorization: `Bearer ${token}`, 'content-type': 'application/json' },
    body: JSON.stringify({
      tool: 'set_secret',
      args: {
        participantId,
        key: `teams-bot/${appId}`,
        secret: JSON.stringify({ appId, clientSecret, tenantId }),
      },
    }),
  });
  if (!response.ok) throw new Error(`set_secret failed: ${response.status}`);
  console.log(`Install: teams app get ${appId} --install-link`);
}

console.log(JSON.stringify({ participantId, appId, endpoint, outDir }, null, 2));
```

```js
#!/usr/bin/env node
// scripts/provision-all.mjs
import { execFileSync } from 'node:child_process';
import { readFile, readdir } from 'node:fs/promises';
import { join, resolve } from 'node:path';

const workspace = resolve(process.env.LEGION_WORKSPACE ?? process.cwd());
const baseUrl = process.argv[process.argv.indexOf('--base-url') + 1];
if (!baseUrl) {
  console.error('Usage: provision-all.mjs --base-url <url>');
  process.exit(2);
}

const dir = join(workspace, '.legion/collective/participants');
for (const file of await readdir(dir)) {
  if (!file.endsWith('.json')) continue;
  const participant = JSON.parse(await readFile(join(dir, file), 'utf8'));
  if (!participant.identities?.some((i) => i.connector === 'teams-bot')) continue;
  execFileSync('node', [
    join(import.meta.dirname, 'provision-teams-bot.mjs'),
    '--participant', participant.id,
    '--base-url', baseUrl,
    '--workspace', workspace,
  ], { stdio: 'inherit' });
}
```

- [ ] **Step 2: Dry-run against a temp participant**

```bash
mkdir -p /tmp/teams-provision/.legion/collective/participants
cat > /tmp/teams-provision/.legion/collective/participants/assistant.json <<'JSON'
{ "id": "assistant", "name": "Assistant", "type": "agent", "tools": {} }
JSON
node scripts/provision-teams-bot.mjs --participant assistant --base-url https://teams.example.com --workspace /tmp/teams-provision --dry-run
```

Expected: prints JSON with `appId: "<app-id>"`, writes `manifest.json` + icons under `/tmp/teams-provision/.legion/teams-apps/assistant`, and adds a `teams-bot` identity with `externalId: "<app-id>"`.

- [ ] **Step 3: Commit**

```bash
git add scripts/provision-teams-bot.mjs scripts/provision-all.mjs
git commit -m "feat: provision per-participant Teams bots"
```

---

## Task 14: Documentation and deployment config

**Files:**
- Create: `README.md`, `docs/deploy.md`

- [ ] **Step 1: Write `README.md`**

Include:

- What the connector does (one bot per participant, personal 1:1).
- Prerequisites: Node 22, an M365 tenant with Teams and custom app upload enabled, a public HTTPS endpoint or Socket Mode.
- Install: `npm install`, `npm run build`.
- Config example (from the spec).
- Bot credentials: `teams-bot/<appId>` in the SecretStore; `LEGION_SECRETS_KEY` required.
- Identity: `teams-bot` on agents, `teams` on users; `/pair` uses the core `pair` tool.
- Commands: `/new`, `/pair`, `/whoami`, `/help`.
- Transport: HTTP default, Socket Mode opt-in.
- Provisioning: `scripts/provision-teams-bot.mjs` / `provision-all.mjs`; install each bot in Teams.
- Testing: `npm test`, `LEGION_TEAMS_INTEGRATION=1`.

- [ ] **Step 2: Write `docs/deploy.md`**

Include:

- Docker mount: connector `dist/` at `/app/legion-connector-teams`.
- `.legion/config.json` entry.
- Reverse proxy/tunnel mapping `teams.prograde-labs.com` → `legion-collective:3978`.
- Socket Mode alternative (no port).
- SecretStore setup (`LEGION_SECRETS_KEY`).
- Tenant notes: Business Basic trial expiry, sideloading, per-bot install.
- Logs to watch: `[legion] external connector 'teams' started`, per-bot start/stop, 403 on uninstalled bots.

- [ ] **Step 3: Commit**

```bash
git add README.md docs/deploy.md
git commit -m "docs: document Teams connector setup and deployment"
```

---

## Task 15: Integration test and repository remotes

**Files:**
- Create: `src/integration.test.ts`
- Modify: `/home/chris/workspace/repos.yaml`

- [ ] **Step 1: Add the opt-in integration test**

```ts
// src/integration.test.ts
import { describe, expect, it } from 'vitest';

const enabled = process.env.LEGION_TEAMS_INTEGRATION === '1';

describe.skipIf(!enabled)('live Teams integration', () => {
  it('starts a Socket Mode bot and receives an install activity', async () => {
    const { App } = await import('@microsoft/teams.apps');
    const app = new App({
      clientId: process.env.TEAMS_CLIENT_ID,
      clientSecret: process.env.TEAMS_CLIENT_SECRET,
      tenantId: process.env.TEAMS_TENANT_ID,
      socketMode: { geos: ['amer'] },
    });
    const seen: string[] = [];
    app.on('install.add', async ({ activity }) => { seen.push(activity.conversation?.id ?? ''); });
    await app.start();
    try {
      await new Promise((resolve) => setTimeout(resolve, 5_000));
      expect(Array.isArray(seen)).toBe(true);
    } finally {
      await app.stop();
    }
  }, 30_000);
});
```

Run (optional, with a live tenant):

```bash
TEAMS_CLIENT_ID=... TEAMS_CLIENT_SECRET=... TEAMS_TENANT_ID=... LEGION_TEAMS_INTEGRATION=1 npm test -- src/integration.test.ts
```

- [ ] **Step 2: Add the `repos.yaml` entry**

In `/home/chris/workspace/repos.yaml`, after the `legion-connector-telegram-v2` entry:

```yaml
  - name: legion-connector-teams
    origin: github
    origin_url: https://github.com/Prograde-Labs/legion-connector-teams.git
    forgejo_role: pull-mirror
    push_mirror: null
    home_branch: main
    workflow: trunk
    role: active
    threads: [2026-11-21-teams-connector]
```

- [ ] **Step 3: Create the GitHub repo and push**

```bash
cd /home/chris/workspace/repos/legion-connector-teams
gh repo create Prograde-Labs/legion-connector-teams --private --source=. --remote=github --push
```

Expected: the repo exists on GitHub and `git remote -v` shows `github`.

- [ ] **Step 4: Commit the workspace bookkeeping**

```bash
cd /home/chris/workspace
git add repos.yaml
git commit -m "chore: register legion-connector-teams"
```

---

## Self-review notes

- **Spec coverage:** Tasks 1 (core `pair`), 2-3 (package/state), 4 (credentials/directory), 5 (identity/pairing), 6 (commands), 7 (render), 8 (communicate), 9 (approvals), 10 (inbound), 11 (outbound), 12 (transport/pool), 13 (provisioning), 14 (deployment), 15 (integration/remotes).
- **Deferred by design:** group chats, channels, files, OAuth/Graph, long-run streaming handoff, feedback/citations.
- **Known limitation:** `StreamChunk` has no per-tool result chunk, so tool cards show invocation and args only.
- **Known limitation:** Teams streaming is 1:1 and two-minute bounded; the SDK finalizes with buffered text on timeout.

