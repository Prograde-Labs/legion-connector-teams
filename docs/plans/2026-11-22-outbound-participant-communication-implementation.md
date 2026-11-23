# Outbound Participant Communication Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Let a bot-owning Legion participant initiate conversations with existing Teams users and groups, with truthful delivery outcomes and reply-thread continuity, through a connector-agnostic core seam.

**Architecture:** Add `Connector.canDeliver(participantId)` and make `Connector.deliver()` return a `DeliveryOutcome`. `UserDeliveryRuntime` asks `ConnectorRegistry.getDeliverableConnectors()` and aggregates outcomes. `MessageRouter` maps a failed outcome to `status: 'error'`, which `communicate` already surfaces. Teams implements the contract first; Web and Telegram adopt it.

**Tech Stack:** TypeScript, Node >= 22.12.0, Vitest, Microsoft Teams SDK 2.2.1, Legion core/runtime monorepo, external connector packages.

**Spec:** `repos/legion-connector-teams/docs/specs/2026-11-22-outbound-participant-communication-design.md`

## Global Constraints

- Node >= 22.12.0; Teams SDK packages are pinned to 2.2.1.
- No connector state schema change; reuse `legionConversations` and `legionConversationTargets`.
- Sender must own the platform bot. The live `assistant` is the only current owner of the Teams bot.
- Only existing Legion participants are valid targets. No roster, cold outreach, or external-ID resolution this round.
- `learnOutboundThread` defaults to `true`.
- No outbox, retry queue, delivery receipts, web UI participant picker, Teams channels, or per-human group participants.
- Existing inbound, approvals, typing, streaming, and personal/group behavior must stay green.
- `canDeliver` is a cheap identity/liveness predicate, never an I/O probe. `deliver` remains the source of truth.
- Do not put credentials or secrets in delivery reasons. Use `describeError` where available.

---

### Task 0: Create the cross-repo worktrees and verify the baseline

**Files:**
- No tracked files. Creates worktrees and builds dependencies.

**Interfaces:**
- Consumes: existing `scripts/thread.sh` convention from `/home/chris/workspace/AGENTS.md`.
- Produces: `threads/2026-11-22-outbound-communication/legion` and `threads/2026-11-22-outbound-communication/legion-connector-teams`, both on `feat/outbound-communication`, with the Teams package resolving the thread's Legion core.

- [ ] **Step 1: Create the feature branches and worktrees**

Run from `/home/chris/workspace`:

```bash
git -C repos/legion branch feat/outbound-communication main
git -C repos/legion-connector-teams branch feat/outbound-communication main
scripts/thread.sh outbound-communication legion feat/outbound-communication
scripts/thread.sh outbound-communication legion-connector-teams feat/outbound-communication
```

Expected: `threads/2026-11-22-outbound-communication/legion` and `threads/2026-11-22-outbound-communication/legion-connector-teams` exist. The Teams branch includes spec commit `f824c18`.

- [ ] **Step 2: Install dependencies in both worktrees**

```bash
cd threads/2026-11-22-outbound-communication/legion
npm install
cd ../legion-connector-teams
npm install
```

If `npm install` cannot reach the network, copy the existing installs instead:

```bash
cp -a /home/chris/workspace/repos/legion/node_modules threads/2026-11-22-outbound-communication/legion/
cp -a /home/chris/workspace/repos/legion-connector-teams/node_modules threads/2026-11-22-outbound-communication/legion-connector-teams/
```

The Teams `@legion-collective/core` symlink is relative (`../../../legion/packages/core`), so from the thread worktree it resolves to `threads/2026-11-22-outbound-communication/legion/packages/core`.

- [ ] **Step 3: Build Legion so Teams can resolve core types**

```bash
cd threads/2026-11-22-outbound-communication/legion
npm run build
```

Expected: `packages/core/dist/index.d.ts` exists and includes the current connector exports.

- [ ] **Step 4: Verify both baselines**

```bash
cd threads/2026-11-22-outbound-communication/legion
npm test
npm run typecheck
cd ../legion-connector-teams
npm test
npm run typecheck
```

Expected: Legion `1345 passed / 21 skipped`; Teams `204 passed / 1 skipped`; both typechecks clean. Do not proceed if either baseline is red.

- [ ] **Step 5: Commit**

No commit for this task. The worktrees are untracked setup.

---

### Task 1: Add the core delivery contract and registry predicate

**Files:**
- Modify: `packages/core/src/connectors/Connector.ts`
- Modify: `packages/core/src/connectors/ConnectorRegistry.ts`
- Modify: `packages/core/src/connectors/index.ts`
- Test: `packages/core/src/connectors/ConnectorRegistry.test.ts`

**Interfaces:**
- Consumes: nothing new.
- Produces:
  - `export type DeliveryStatus = 'delivered' | 'accepted' | 'not-reachable' | 'failed'`
  - `export interface DeliveryOutcome { status: DeliveryStatus; reason?: string }`
  - `Connector.canDeliver?(participantId: string): boolean`
  - `Connector.deliver(message): Promise<DeliveryOutcome>`
  - `ConnectorRegistry.getDeliverableConnectors(participantId: string): Connector[]`

- [ ] **Step 1: Write the failing registry tests**

In `packages/core/src/connectors/ConnectorRegistry.test.ts`, replace the `makeConnector` helper with one that accepts an optional predicate and returns a delivered outcome:

```ts
function makeConnector(
  name: string,
  canDeliver?: (participantId: string) => boolean,
): Connector {
  return {
    name,
    ...(canDeliver ? { canDeliver } : {}),
    async start() {},
    async deliver() {
      return { status: 'delivered' };
    },
    async stop() {},
  };
}
```

Add these tests before the closing `});` of the `describe('ConnectorRegistry')` block:

```ts
it('getDeliverableConnectors includes connectors whose predicate accepts the participant', () => {
  const reg = new ConnectorRegistry();
  reg.register(makeConnector('teams', (id) => id === 'user-1'));
  reg.register(makeConnector('web', (id) => id === 'user-2'));
  expect(reg.getDeliverableConnectors('user-1').map((c) => c.name)).toEqual(['teams']);
});

it('getDeliverableConnectors includes connectors without a predicate', () => {
  const reg = new ConnectorRegistry();
  reg.register(makeConnector('legacy'));
  expect(reg.getDeliverableConnectors('user-1').map((c) => c.name)).toEqual(['legacy']);
});

it('getDeliverableConnectors does not depend on active state', () => {
  const reg = new ConnectorRegistry();
  reg.register(makeConnector('teams', () => true));
  expect(reg.getActiveConnectors('user-1')).toHaveLength(0);
  expect(reg.getDeliverableConnectors('user-1')).toHaveLength(1);
});
```

- [ ] **Step 2: Run the tests to verify they fail**

```bash
cd threads/2026-11-22-outbound-communication/legion
npx vitest run packages/core/src/connectors/ConnectorRegistry.test.ts
```

Expected: FAIL with `reg.getDeliverableConnectors is not a function`.

- [ ] **Step 3: Add the contract types**

In `packages/core/src/connectors/Connector.ts`, add above the `Connector` interface:

```ts
/** Result of one connector's attempt to push a message outbound. */
export type DeliveryStatus = 'delivered' | 'accepted' | 'not-reachable' | 'failed';

export interface DeliveryOutcome {
  status: DeliveryStatus;
  /** Human-readable reason for `not-reachable` or `failed`. */
  reason?: string;
}
```

In the `Connector` interface, add the predicate and change `deliver`:

```ts
export interface Connector {
  /** Unique connector name, e.g. `'web'`, `'teams'`, `'slack'`. */
  readonly name: string;

  /**
   * Whether this connector can deliver to the participant. Optional for
   * compatibility; an absent predicate means "ask me". Implementations should
   * base this on the participant's identities and their own liveness, not on
   * an I/O probe. `deliver` remains the source of truth for reachability.
   */
  canDeliver?(participantId: string): boolean;

  /** Called once at process startup to hand over the `ConnectorContext`. */
  start(ctx: ConnectorContext): Promise<void>;

  /**
   * Push a message outbound to the external entity the connector fronts.
   * Returns whether the platform accepted it, or why it could not be reached.
   */
  deliver(message: {
    id: string;
    conversationId: string;
    senderId: string;
    recipientId: string;
    replyTo?: string;
    content: string;
    timestamp: string;
  }): Promise<DeliveryOutcome>;

  /** Called on process shutdown. Release all resources. */
  stop(): Promise<void>;
}
```

In `packages/core/src/connectors/index.ts`, export the new types:

```ts
export type {
  Connector,
  ConnectorContext,
  ConnectorSubmitOptions,
  DeliveryOutcome,
  DeliveryStatus,
} from './Connector.js';
```

- [ ] **Step 4: Add the registry method**

In `packages/core/src/connectors/ConnectorRegistry.ts`, update the class doc comment and add:

```ts
  /**
   * Returns every registered connector whose `canDeliver` predicate accepts
   * `participantId`. Connectors without a predicate are treated as reachable.
   * This is independent of `setActive`, which remains presence-only.
   */
  getDeliverableConnectors(participantId: string): Connector[] {
    return this.getAll().filter(
      (connector) => connector.canDeliver?.(participantId) ?? true,
    );
  }
```

Update the class comment from `getActiveConnectors() — called by UserDeliveryRuntime` to say `getActiveConnectors()` is presence-only and `getDeliverableConnectors()` is what `UserDeliveryRuntime` uses.

- [ ] **Step 5: Run the tests to verify they pass**

```bash
npx vitest run packages/core/src/connectors/ConnectorRegistry.test.ts
```

Expected: PASS.

- [ ] **Step 6: Commit**

```bash
git add packages/core/src/connectors/Connector.ts \
        packages/core/src/connectors/ConnectorRegistry.ts \
        packages/core/src/connectors/index.ts \
        packages/core/src/connectors/ConnectorRegistry.test.ts
git commit -m "feat(core): add connector delivery outcomes"
```

---

### Task 2: Return delivery outcomes from the web connector

**Files:**
- Modify: `packages/runtime/src/server/WebConnector.ts`
- Test: `packages/runtime/src/server/WebConnector.test.ts`

**Interfaces:**
- Consumes: `DeliveryOutcome` from Task 1.
- Produces:
  - `WebConnector.canDeliver(participantId: string): boolean`
  - `WebConnector.deliver(message): Promise<DeliveryOutcome>`

- [ ] **Step 1: Update the web tests to expect outcomes**

In `packages/runtime/src/server/WebConnector.test.ts`, add a `web` identity to the bootstrap operator in `makeCollective()`:

```ts
  await storage.writeJson('collective/participants/op-1.json', {
    id: 'op-1',
    name: 'Operator',
    type: 'user',
    tools: { list_participants: 'auto' },
    operator: true,
    protected: true,
    status: 'active',
    identities: [{ connector: 'web', externalId: 'op-1' }],
  });
```

Replace the three assertions in `describe('WebConnector: deliver()')` with outcome assertions, and add two tests:

```ts
describe('WebConnector: deliver()', () => {
  it('reports accepted when the recipient has no active connections', async () => {
    const { connector } = await makeConnector();
    await expect(
      connector.deliver({
        id: 'm1',
        conversationId: 'c1',
        senderId: 'agent-1',
        recipientId: 'op-1',
        content: 'hello',
        timestamp: new Date().toISOString(),
      }),
    ).resolves.toEqual({ status: 'accepted' });
    await connector.stop();
  });

  it('reports delivered when it sends to an open socket', async () => {
    const { connector } = await makeConnector();
    const sent: string[] = [];
    const fakeSocket = {
      readyState: 1,
      send: (data: string) => {
        sent.push(data);
      },
    };
    (connector as any).connections.set('op-1', new Set([fakeSocket]));

    await expect(
      connector.deliver({
        id: 'm1',
        conversationId: 'c1',
        senderId: 'agent-1',
        recipientId: 'op-1',
        content: 'hi',
        timestamp: new Date().toISOString(),
      }),
    ).resolves.toEqual({ status: 'delivered' });

    expect(sent).toHaveLength(1);
    const msg = JSON.parse(sent[0]) as { type: string };
    expect(msg.type).toBe('message');
    await connector.stop();
  });

  it('reports accepted when every socket is closed', async () => {
    const { connector } = await makeConnector();
    const sent: string[] = [];
    const closedSocket = {
      readyState: 3,
      send: (data: string) => {
        sent.push(data);
      },
    };
    (connector as any).connections.set('op-1', new Set([closedSocket]));

    await expect(
      connector.deliver({
        id: 'm1',
        conversationId: 'c1',
        senderId: 'agent-1',
        recipientId: 'op-1',
        content: 'hi',
        timestamp: new Date().toISOString(),
      }),
    ).resolves.toEqual({ status: 'accepted' });
    expect(sent).toHaveLength(0);
    await connector.stop();
  });

  it('reports failed when a socket send throws', async () => {
    const { connector } = await makeConnector();
    const throwingSocket = {
      readyState: 1,
      send: () => {
        throw new Error('socket closed');
      },
    };
    (connector as any).connections.set('op-1', new Set([throwingSocket]));

    await expect(
      connector.deliver({
        id: 'm1',
        conversationId: 'c1',
        senderId: 'agent-1',
        recipientId: 'op-1',
        content: 'hi',
        timestamp: new Date().toISOString(),
      }),
    ).resolves.toEqual({ status: 'failed', reason: 'socket closed' });
    await connector.stop();
  });

  it('reports deliverability from the web identity', async () => {
    const { connector } = await makeConnector();
    expect(connector.canDeliver('op-1')).toBe(true);
    expect(connector.canDeliver('nobody')).toBe(false);
    await connector.stop();
  });
});
```

- [ ] **Step 2: Run the tests to verify they fail**

```bash
cd threads/2026-11-22-outbound-communication/legion
npx vitest run packages/runtime/src/server/WebConnector.test.ts
```

Expected: FAIL because `deliver()` resolves to `undefined` and `canDeliver` is not defined.

- [ ] **Step 3: Implement the web contract**

In `packages/runtime/src/server/WebConnector.ts`, add `DeliveryOutcome` to the core type import:

```ts
import type { Connector, ConnectorContext, DeliveryOutcome } from '@legion-collective/core';
```

Add `canDeliver` before `deliver`:

```ts
  canDeliver(participantId: string): boolean {
    const participant = this.deps.collective.get(participantId);
    return Boolean(
      participant?.identities?.some((identity) => identity.connector === 'web'),
    );
  }
```

Replace `deliver` with:

```ts
  async deliver(message: {
    id: string;
    conversationId: string;
    senderId: string;
    recipientId: string;
    replyTo?: string;
    content: string;
    timestamp: string;
  }): Promise<DeliveryOutcome> {
    const sockets = this.connections.get(message.recipientId);
    if (!sockets?.size) return { status: 'accepted' };

    const payload = JSON.stringify({ type: 'message', data: message });
    let delivered = false;
    let failure: string | undefined;

    for (const socket of sockets) {
      try {
        if ((socket as any).readyState === WebSocket.OPEN) {
          socket.send(payload);
          delivered = true;
        }
      } catch (error) {
        failure = error instanceof Error ? error.message : String(error);
      }
    }

    if (delivered) return { status: 'delivered' };
    if (failure) return { status: 'failed', reason: failure };
    return { status: 'accepted' };
  }
```

- [ ] **Step 4: Run the tests to verify they pass**

```bash
npx vitest run packages/runtime/src/server/WebConnector.test.ts
```

Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add packages/runtime/src/server/WebConnector.ts \
        packages/runtime/src/server/WebConnector.test.ts
git commit -m "feat(web): return delivery outcomes"
```

---

### Task 3: Deliver through deliverable connectors and aggregate outcomes

**Files:**
- Modify: `packages/core/src/runtime/Runtime.ts`
- Modify: `packages/core/src/runtime/UserDeliveryRuntime.ts`
- Test: `packages/core/src/runtime/UserDeliveryRuntime.test.ts`

**Interfaces:**
- Consumes: `DeliveryOutcome`, `ConnectorRegistry.getDeliverableConnectors` from Task 1.
- Produces:
  - `RuntimeResult` void variant gains `delivery?: DeliveryOutcome`.
  - `UserDeliveryRuntime.handle()` returns `{ kind: 'void', delivery }` for user deliveries.

- [ ] **Step 1: Rewrite the runtime tests**

Replace `packages/core/src/runtime/UserDeliveryRuntime.test.ts` with:

```ts
import { UserDeliveryRuntime } from './UserDeliveryRuntime.js';
import { ConnectorRegistry } from '../connectors/ConnectorRegistry.js';
import type { RuntimeContext } from './Runtime.js';
import type { MessageData } from '@legion-collective/types';
import type { Connector, DeliveryOutcome } from '../connectors/Connector.js';

function makeConnector(
  name: string,
  outcome: DeliveryOutcome = { status: 'delivered' },
  canDeliver?: (participantId: string) => boolean,
): Connector & { delivered: unknown[] } {
  return {
    name,
    delivered: [],
    ...(canDeliver ? { canDeliver } : {}),
    async start() {},
    async deliver(msg) {
      this.delivered.push(msg);
      return outcome;
    },
    async stop() {},
  };
}

const inbound: MessageData = {
  id: 'm1',
  parentId: null,
  conversationId: 'c1',
  senderId: 'op',
  recipientId: 'user-a',
  role: 'user',
  content: 'hello',
  status: 'active',
  timestamp: '2026-01-01T00:00:00.000Z',
};

function makeCtx(): RuntimeContext {
  const emitted: unknown[] = [];
  return {
    participant: { id: 'user-a', name: 'Alice', type: 'user', tools: {} },
    conversationId: 'c1',
    eventBus: { emit: (_e: string, p: unknown) => emitted.push(p), on: () => () => {} } as any,
    _emitted: emitted,
  } as any;
}

describe('UserDeliveryRuntime', () => {
  it('returns not-reachable when no registry is provided', async () => {
    const runtime = new UserDeliveryRuntime('user-a');
    await expect(runtime.handle(inbound, makeCtx())).resolves.toEqual({
      kind: 'void',
      delivery: { status: 'not-reachable', reason: 'no connector can reach user-a' },
    });
  });

  it('emits message:delivered regardless of connector presence', async () => {
    const ctx = makeCtx();
    const runtime = new UserDeliveryRuntime('user-a');
    await runtime.handle(inbound, ctx);
    expect((ctx as any)._emitted).toHaveLength(1);
  });

  it('calls deliver() on deliverable connectors without setActive', async () => {
    const registry = new ConnectorRegistry();
    const conn = makeConnector('web');
    registry.register(conn);
    const runtime = new UserDeliveryRuntime('user-a', registry);
    await runtime.handle(inbound, makeCtx());
    expect(conn.delivered).toHaveLength(1);
    expect((conn.delivered[0] as any).content).toBe('hello');
  });

  it('excludes connectors whose canDeliver predicate rejects the participant', async () => {
    const registry = new ConnectorRegistry();
    const accepted = makeConnector('teams', { status: 'delivered' }, (id) => id === 'user-a');
    const rejected = makeConnector('web', { status: 'delivered' }, () => false);
    registry.register(accepted);
    registry.register(rejected);
    const runtime = new UserDeliveryRuntime('user-a', registry);
    await runtime.handle(inbound, makeCtx());
    expect(accepted.delivered).toHaveLength(1);
    expect(rejected.delivered).toHaveLength(0);
  });

  it('fans out to every deliverable connector and reports the first success', async () => {
    const registry = new ConnectorRegistry();
    const first = makeConnector('teams', { status: 'delivered' });
    const second = makeConnector('web', { status: 'accepted' });
    registry.register(first);
    registry.register(second);
    const runtime = new UserDeliveryRuntime('user-a', registry);
    await expect(runtime.handle(inbound, makeCtx())).resolves.toEqual({
      kind: 'void',
      delivery: { status: 'delivered' },
    });
    expect(first.delivered).toHaveLength(1);
    expect(second.delivered).toHaveLength(1);
  });

  it('reports accepted as a success', async () => {
    const registry = new ConnectorRegistry();
    registry.register(makeConnector('web', { status: 'accepted' }));
    const runtime = new UserDeliveryRuntime('user-a', registry);
    await expect(runtime.handle(inbound, makeCtx())).resolves.toEqual({
      kind: 'void',
      delivery: { status: 'accepted' },
    });
  });

  it('reports failed when every connector fails or is unreachable', async () => {
    const registry = new ConnectorRegistry();
    registry.register(makeConnector('teams', { status: 'failed', reason: '403' }));
    registry.register(makeConnector('web', { status: 'not-reachable', reason: 'offline' }));
    const runtime = new UserDeliveryRuntime('user-a', registry);
    await expect(runtime.handle(inbound, makeCtx())).resolves.toEqual({
      kind: 'void',
      delivery: { status: 'failed', reason: '403; offline' },
    });
  });

  it('reports not-reachable when every connector is unreachable', async () => {
    const registry = new ConnectorRegistry();
    registry.register(makeConnector('teams', { status: 'not-reachable', reason: 'no bot' }));
    registry.register(makeConnector('web', { status: 'not-reachable', reason: 'no socket' }));
    const runtime = new UserDeliveryRuntime('user-a', registry);
    await expect(runtime.handle(inbound, makeCtx())).resolves.toEqual({
      kind: 'void',
      delivery: { status: 'not-reachable', reason: 'no bot; no socket' },
    });
  });

  it('converts thrown connector errors to failed outcomes and continues', async () => {
    const registry = new ConnectorRegistry();
    const broken: Connector = {
      name: 'broken',
      async start() {},
      async stop() {},
      async deliver() {
        throw new Error('network error');
      },
    };
    const good = makeConnector('good');
    registry.register(broken);
    registry.register(good);
    const runtime = new UserDeliveryRuntime('user-a', registry);
    await expect(runtime.handle(inbound, makeCtx())).resolves.toEqual({
      kind: 'void',
      delivery: { status: 'delivered' },
    });
    expect(good.delivered).toHaveLength(1);
  });

  it('reports failed when every connector throws', async () => {
    const registry = new ConnectorRegistry();
    registry.register({
      name: 'broken',
      async start() {},
      async stop() {},
      async deliver() {
        throw new Error('network error');
      },
    });
    const runtime = new UserDeliveryRuntime('user-a', registry);
    await expect(runtime.handle(inbound, makeCtx())).resolves.toEqual({
      kind: 'void',
      delivery: { status: 'failed', reason: 'network error' },
    });
  });
});
```

- [ ] **Step 2: Run the tests to verify they fail**

```bash
npx vitest run packages/core/src/runtime/UserDeliveryRuntime.test.ts
```

Expected: FAIL because the runtime still uses `getActiveConnectors` and returns a plain `void`.

- [ ] **Step 3: Extend `RuntimeResult`**

In `packages/core/src/runtime/Runtime.ts`, add the type import at the top:

```ts
import type { DeliveryOutcome } from '../connectors/Connector.js';
```

Change the `void` variant:

```ts
  | {
      kind: 'void';
      delivery?: DeliveryOutcome;
      actions?: MiddlewareActionResult[];
    };
```

- [ ] **Step 4: Rewrite `UserDeliveryRuntime`**

Replace the file body with:

```ts
import type { MessageData } from '@legion-collective/types';
import type { Runtime, RuntimeContext, RuntimeResult } from './Runtime.js';
import type { ConnectorRegistry } from '../connectors/ConnectorRegistry.js';
import type { DeliveryOutcome } from '../connectors/Connector.js';

export class UserDeliveryRuntime implements Runtime {
  constructor(
    private participantId: string,
    private connectorRegistry?: ConnectorRegistry,
  ) {}

  async handle(incoming: MessageData, context: RuntimeContext): Promise<RuntimeResult> {
    const connectors = this.connectorRegistry?.getDeliverableConnectors(this.participantId) ?? [];

    const outcomes: DeliveryOutcome[] = [];
    if (connectors.length === 0) {
      outcomes.push({
        status: 'not-reachable',
        reason: `no connector can reach ${this.participantId}`,
      });
    } else {
      for (const connector of connectors) {
        try {
          outcomes.push(
            await connector.deliver({
              id: incoming.id,
              conversationId: incoming.conversationId,
              senderId: incoming.senderId,
              recipientId: incoming.recipientId,
              replyTo: incoming.replyTo,
              content: incoming.content,
              timestamp: incoming.timestamp,
            }),
          );
        } catch (error) {
          outcomes.push({
            status: 'failed',
            reason: error instanceof Error ? error.message : String(error),
          });
        }
      }
    }

    context.eventBus.emit('message:delivered', {
      conversationId: context.conversationId,
      recipientId: this.participantId,
      messageId: incoming.id,
    });

    const successful = outcomes.find(
      (outcome) => outcome.status === 'delivered' || outcome.status === 'accepted',
    );
    if (successful) return { kind: 'void', delivery: successful };

    const failed = outcomes.find((outcome) => outcome.status === 'failed');
    const reason = outcomes.map((outcome) => outcome.reason ?? outcome.status).join('; ');
    return {
      kind: 'void',
      delivery: {
        status: failed ? 'failed' : 'not-reachable',
        reason,
      },
    };
  }
}
```

- [ ] **Step 5: Run the tests to verify they pass**

```bash
npx vitest run packages/core/src/runtime/UserDeliveryRuntime.test.ts
```

Expected: PASS.

- [ ] **Step 6: Commit**

```bash
git add packages/core/src/runtime/Runtime.ts \
        packages/core/src/runtime/UserDeliveryRuntime.ts \
        packages/core/src/runtime/UserDeliveryRuntime.test.ts
git commit -m "feat(core): deliver through deliverable connectors"
```

---

### Task 4: Surface failed delivery as an error in `MessageRouter`

**Files:**
- Modify: `packages/core/src/runtime/MessageRouter.ts`
- Test: `packages/core/src/runtime/MessageRouter.test.ts`

**Interfaces:**
- Consumes: `RuntimeResult.void.delivery` from Task 3.
- Produces: `MessageRouter.completeResponse` maps `failed` and `not-reachable` to `{ status: 'error', error: reason }`.

- [ ] **Step 1: Write the failing router tests**

At the end of `packages/core/src/runtime/MessageRouter.test.ts`, add a new describe block. Add `import type { RuntimeResult } from './Runtime.js';` to the existing imports.

```ts
describe('MessageRouter: delivery outcomes', () => {
  let dir: string;
  beforeEach(async () => {
    dir = await mkdtemp(join(tmpdir(), 'legion-router-delivery-'));
  });
  afterEach(async () => {
    await rm(dir, { recursive: true, force: true });
  });

  async function routerFor(outcome: RuntimeResult) {
    const { baseContext, store, eventBus, collective } = await setup(dir);
    const registry = new RuntimeRegistry();
    registry.registerFactory('mock', () => ({
      async handle() {
        return outcome;
      },
    }));
    return {
      router: new MessageRouter(store, registry, collective, eventBus),
      baseContext,
    };
  }

  it('maps a failed delivery outcome to an error', async () => {
    const { router, baseContext } = await routerFor({
      kind: 'void',
      delivery: { status: 'failed', reason: 'no live Teams bot' },
    });
    const result = await router.send({
      senderId: 'op',
      recipientId: 'mock-1',
      message: 'hello',
      context: baseContext,
    });
    expect(result).toMatchObject({ status: 'error', error: 'no live Teams bot' });
  });

  it('maps a not-reachable delivery outcome to an error', async () => {
    const { router, baseContext } = await routerFor({
      kind: 'void',
      delivery: { status: 'not-reachable', reason: 'no connector can reach user-a' },
    });
    const result = await router.send({
      senderId: 'op',
      recipientId: 'mock-1',
      message: 'hello',
      context: baseContext,
    });
    expect(result).toMatchObject({
      status: 'error',
      error: 'no connector can reach user-a',
    });
  });

  it('keeps delivered and accepted outcomes as success', async () => {
    for (const status of ['delivered', 'accepted'] as const) {
      const { router, baseContext } = await routerFor({
        kind: 'void',
        delivery: { status },
      });
      const result = await router.send({
        senderId: 'op',
        recipientId: 'mock-1',
        message: 'hello',
        context: baseContext,
      });
      expect(result.status).toBe('success');
    }
  });

  it('keeps a void result without a delivery outcome as success', async () => {
    const { router, baseContext } = await routerFor({ kind: 'void' });
    const result = await router.send({
      senderId: 'op',
      recipientId: 'mock-1',
      message: 'hello',
      context: baseContext,
    });
    expect(result.status).toBe('success');
  });
});
```

- [ ] **Step 2: Run the tests to verify they fail**

```bash
npx vitest run packages/core/src/runtime/MessageRouter.test.ts -t "delivery outcomes"
```

Expected: FAIL because the failed outcome still returns `status: 'success'`.

- [ ] **Step 3: Implement the mapping**

In `packages/core/src/runtime/MessageRouter.ts`, replace the final `// kind === 'void'` block of `completeResponse`:

```ts
    // kind === 'void'
    if (
      result.delivery &&
      result.delivery.status !== 'delivered' &&
      result.delivery.status !== 'accepted'
    ) {
      return {
        conversationId: operation.thread.id,
        status: 'error',
        error: result.delivery.reason ?? result.delivery.status,
      };
    }
    return { conversationId: operation.thread.id, status: 'success' };
```

- [ ] **Step 4: Run the tests to verify they pass**

```bash
npx vitest run packages/core/src/runtime/MessageRouter.test.ts -t "delivery outcomes"
```

Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add packages/core/src/runtime/MessageRouter.ts \
        packages/core/src/runtime/MessageRouter.test.ts
git commit -m "fix(core): surface failed delivery as an error"
```

---

### Task 5: Add the Teams `learnOutboundThread` option and `canDeliver`

**Files:**
- Modify: `repos/legion-connector-teams/src/options.ts`
- Modify: `repos/legion-connector-teams/src/connector.ts`
- Test: `repos/legion-connector-teams/src/options.test.ts`
- Test: `repos/legion-connector-teams/src/connector.test.ts`

**Interfaces:**
- Consumes: `DeliveryOutcome` and `Connector.canDeliver` from Task 1.
- Produces:
  - `TeamsConnectorOptions.learnOutboundThread?: boolean`
  - `ResolvedOptions.learnOutboundThread: boolean` (default `true`)
  - `TeamsConnector.canDeliver(participantId: string): boolean` (true for `teams` or `teams-group` identities)

- [ ] **Step 1: Write the failing option tests**

Add to `src/options.test.ts`:

```ts
describe('resolveOptions outbound thread learning', () => {
  it('defaults learnOutboundThread on', () => {
    expect(resolveOptions({ stateDir: '/tmp/x' }).learnOutboundThread).toBe(true);
  });

  it('can disable learnOutboundThread', () => {
    expect(resolveOptions({ stateDir: '/tmp/x', learnOutboundThread: false }).learnOutboundThread)
      .toBe(false);
  });
});
```

- [ ] **Step 2: Run the option tests to verify they fail**

```bash
cd threads/2026-11-22-outbound-communication/legion-connector-teams
npx vitest run src/options.test.ts
```

Expected: FAIL because `learnOutboundThread` is `undefined`.

- [ ] **Step 3: Add the option**

In `src/options.ts`, add to `TeamsConnectorOptions` after `groupSenderPrefix`:

```ts
  learnOutboundThread?: boolean;
```

Add to `ResolvedOptions` after `groupSenderPrefix`:

```ts
  learnOutboundThread: boolean;
```

In `resolveOptions`, add after `groupSenderPrefix`:

```ts
    learnOutboundThread: options.learnOutboundThread ?? true,
```

- [ ] **Step 4: Run the option tests to verify they pass**

```bash
npx vitest run src/options.test.ts
```

Expected: PASS.

- [ ] **Step 5: Write the failing `canDeliver` test**

In `src/connector.test.ts`, add a group participant to `setup()`:

```ts
    { id: 'team-eng', name: 'Team Eng', type: 'user', tools: {}, status: 'active',
      identities: [{ connector: 'teams-group', externalId: '19:g@thread.v2' }] },
```

Add this test inside `describe('TeamsConnector')`:

```ts
  it('reports deliverability from Teams identities', () => {
    const { deps } = setup();
    const connector = new TeamsConnector({ stateDir: '/tmp/teams-can-deliver' }, deps);
    expect(connector.canDeliver('alice')).toBe(true);
    expect(connector.canDeliver('team-eng')).toBe(true);
    expect(connector.canDeliver('assistant')).toBe(false);
    expect(connector.canDeliver('nobody')).toBe(false);
  });
```

- [ ] **Step 6: Run the test to verify it fails**

```bash
npx vitest run src/connector.test.ts -t "reports deliverability"
```

Expected: FAIL because `canDeliver` is not defined.

- [ ] **Step 7: Implement `canDeliver`**

In `src/connector.ts`, add before `async start`:

```ts
  canDeliver(participantId: string): boolean {
    const participant = this.deps.collective.get(participantId);
    return Boolean(
      participant?.identities?.some(
        (identity) =>
          identity.connector === 'teams' || identity.connector === 'teams-group',
      ),
    );
  }
```

- [ ] **Step 8: Run the tests to verify they pass**

```bash
npx vitest run src/options.test.ts src/connector.test.ts
```

Expected: PASS.

- [ ] **Step 9: Commit**

```bash
git add src/options.ts src/connector.ts src/options.test.ts src/connector.test.ts
git commit -m "feat(teams): add outbound deliverability and thread option"
```

---

### Task 6: Return Teams delivery outcomes and learn the outbound thread

**Files:**
- Modify: `repos/legion-connector-teams/src/outbound.ts`
- Modify: `repos/legion-connector-teams/src/connector.ts`
- Test: `repos/legion-connector-teams/src/outbound.test.ts`
- Test: `repos/legion-connector-teams/src/connector.test.ts`

**Interfaces:**
- Consumes: `DeliveryOutcome` from Task 1, `learnOutboundThread` from Task 5.
- Produces:
  - `deliverMessage(deps, message): Promise<DeliveryOutcome>`
  - `TeamsConnector.deliver(message): Promise<DeliveryOutcome>`
  - On success, always `legionConversationTargets[conversationId] = recordKey`; when learning is on, `legionConversations[recordKey] = conversationId`.

- [ ] **Step 1: Update the outbound tests for outcomes and thread learning**

In `src/outbound.test.ts`, update the existing failure tests:

```ts
  it('reports not-reachable when the sender has no Teams bot', async () => {
    const s = setup([
      { id: 'assistant', name: 'Assistant', type: 'agent', tools: {} },
      { id: 'alice', name: 'Alice', type: 'user', tools: {},
        identities: [{ connector: 'teams', externalId: 'aad-1' }] },
    ]);
    await expect(deliverMessage(
      { collective: s.collective, state: s.state as never, stateStore: { save: s.save } as never,
        apps: s.apps as never, options: resolveOptions({ stateDir: '/tmp/x' }) },
      { senderId: 'assistant', recipientId: 'alice', conversationId: 'legion-1', content: 'hello' },
    )).resolves.toEqual({
      status: 'not-reachable',
      reason: "sender 'assistant' has no live Teams bot",
    });
    expect(s.send).not.toHaveBeenCalled();
  });

  it('reports failed and logs a sanitized Teams send error', async () => {
    const s = setup([
      { id: 'assistant', identities: [{ connector: 'teams-bot', externalId: 'app' }] },
    ], { 'app:alice': { conversationId: 'conv-1', updatedAt: 'now' } });
    const error = Object.assign(new Error('403 Bearer abc123'), { authorization: 'Bearer secret-field' });
    s.send.mockRejectedValueOnce(error);
    const log = vi.spyOn(console, 'error').mockImplementation(() => undefined);
    try {
      await expect(deliverMessage(
        { collective: s.collective, state: s.state as never, stateStore: { save: s.save } as never,
          apps: s.apps as never, options: resolveOptions({ stateDir: '/tmp/x' }) },
        { senderId: 'assistant', recipientId: 'alice', conversationId: 'legion-1', content: 'hello' },
      )).resolves.toEqual({
        status: 'failed',
        reason: 'Teams message delivery failed: Error: 403 Bearer [redacted]',
      });
      expect(log).toHaveBeenCalledWith('Teams message delivery failed', 'Error: 403 Bearer [redacted]');
      expect(JSON.stringify(log.mock.calls)).not.toContain('abc123');
      expect(JSON.stringify(log.mock.calls)).not.toContain('secret-field');
    } finally {
      log.mockRestore();
    }
  });

  it('reports failed and logs a sanitized creation error without saving or sending', async () => {
    const s = setup([
      { id: 'assistant', identities: [{ connector: 'teams-bot', externalId: 'app' }] },
      { id: 'alice', identities: [{ connector: 'teams', externalId: 'aad-1' }] },
    ]);
    s.create.mockRejectedValueOnce(new Error('403 Bearer abc123'));
    const log = vi.spyOn(console, 'error').mockImplementation(() => undefined);
    try {
      await expect(deliverMessage(
        { collective: s.collective, state: s.state as never, stateStore: { save: s.save } as never,
          apps: s.apps as never, options: resolveOptions({ stateDir: '/tmp/x' }) },
        { senderId: 'assistant', recipientId: 'alice', conversationId: 'legion-1', content: 'hello' },
      )).resolves.toEqual({
        status: 'failed',
        reason: 'Teams conversation creation failed: Error: 403 Bearer [redacted]',
      });
      expect(log).toHaveBeenCalledWith('Teams conversation creation failed', 'Error: 403 Bearer [redacted]');
      expect(JSON.stringify(log.mock.calls)).not.toContain('abc123');
      expect(s.state.conversations).toEqual({});
      expect(s.save).not.toHaveBeenCalled();
      expect(s.send).not.toHaveBeenCalled();
    } finally {
      log.mockRestore();
    }
  });

  it('reports not-reachable when the bot app or recipient identity is missing', async () => {
    const s = setup([
      { id: 'assistant', identities: [{ connector: 'teams-bot', externalId: 'other' }] },
      { id: 'alice' },
    ]);
    const deps = { collective: s.collective, state: s.state as never, stateStore: { save: s.save } as never,
      apps: s.apps as never, options: resolveOptions({ stateDir: '/tmp/x' }) };
    const message = { senderId: 'assistant', recipientId: 'alice', conversationId: 'legion-1', content: 'hello' };
    await expect(deliverMessage(deps, message)).resolves.toEqual({
      status: 'not-reachable',
      reason: "sender 'assistant' has no live Teams bot",
    });
    expect(s.create).not.toHaveBeenCalled();
    (s.collective.get('assistant')!.identities![0]).externalId = 'app';
    await expect(deliverMessage(deps, message)).resolves.toEqual({
      status: 'not-reachable',
      reason: "no Teams conversation for 'alice'",
    });
    expect(s.create).not.toHaveBeenCalled();
    expect(s.send).not.toHaveBeenCalled();
  });
```

Add these new tests:

```ts
  it('sets the target map and learns the thread by default', async () => {
    const s = setup([
      { id: 'assistant', identities: [{ connector: 'teams-bot', externalId: 'app' }] },
      { id: 'alice', identities: [{ connector: 'teams', externalId: 'aad-1' }] },
    ], { 'app:alice': { conversationId: 'conv-1', updatedAt: 'now' } });
    await expect(deliverMessage(
      { collective: s.collective, state: s.state as never, stateStore: { save: s.save } as never,
        apps: s.apps as never, options: resolveOptions({ stateDir: '/tmp/x' }) },
      { senderId: 'assistant', recipientId: 'alice', conversationId: 'legion-1', content: 'hello' },
    )).resolves.toEqual({ status: 'delivered' });
    expect(s.state.legionConversationTargets['legion-1']).toBe('app:alice');
    expect(s.state.legionConversations['app:alice']).toBe('legion-1');
    expect(s.save).toHaveBeenCalledWith(s.state);
  });

  it('sets the target map without changing the thread when learning is off', async () => {
    const s = setup([
      { id: 'assistant', identities: [{ connector: 'teams-bot', externalId: 'app' }] },
      { id: 'alice', identities: [{ connector: 'teams', externalId: 'aad-1' }] },
    ], { 'app:alice': { conversationId: 'conv-1', updatedAt: 'now' } });
    s.state.legionConversations['app:alice'] = 'legion-previous';
    await expect(deliverMessage(
      { collective: s.collective, state: s.state as never, stateStore: { save: s.save } as never,
        apps: s.apps as never,
        options: resolveOptions({ stateDir: '/tmp/x', learnOutboundThread: false }) },
      { senderId: 'assistant', recipientId: 'alice', conversationId: 'legion-1', content: 'hello' },
    )).resolves.toEqual({ status: 'delivered' });
    expect(s.state.legionConversationTargets['legion-1']).toBe('app:alice');
    expect(s.state.legionConversations['app:alice']).toBe('legion-previous');
    expect(s.save).toHaveBeenCalledWith(s.state);
  });

  it('learns the group thread by default', async () => {
    const s = setup([
      { id: 'assistant', identities: [{ connector: 'teams-bot', externalId: 'app' }] },
      { id: 'team-eng', identities: [{ connector: 'teams-group', externalId: '19:g@thread.v2' }] },
    ]);
    await expect(deliverMessage(
      { collective: s.collective, state: s.state as never, stateStore: { save: s.save } as never,
        apps: s.apps as never, options: resolveOptions({ stateDir: '/tmp/x' }) },
      { senderId: 'assistant', recipientId: 'team-eng', conversationId: 'legion-1', content: 'hello' },
    )).resolves.toEqual({ status: 'delivered' });
    expect(s.state.legionConversationTargets['legion-1']).toBe('app:team-eng:19:g@thread.v2');
    expect(s.state.legionConversations['app:team-eng:19:g@thread.v2']).toBe('legion-1');
  });
```

- [ ] **Step 2: Run the tests to verify they fail**

```bash
npx vitest run src/outbound.test.ts
```

Expected: FAIL because `deliverMessage` still resolves to `undefined` and does not set `legionConversations`.

- [ ] **Step 3: Rewrite `deliverMessage`**

Replace `src/outbound.ts`'s `deliverMessage` (keep the imports and interfaces) with:

```ts
import type { DeliveryOutcome } from '@legion-collective/core';

export async function deliverMessage(
  deps: OutboundDeps,
  message: { senderId: string; recipientId: string; conversationId: string; content: string },
): Promise<DeliveryOutcome> {
  const sender = deps.collective.get(message.senderId);
  const botIdentity = sender?.identities?.find((i) => i.connector === 'teams-bot');
  if (!botIdentity || deps.apps.ownerOf(botIdentity.externalId) !== message.senderId) {
    return {
      status: 'not-reachable',
      reason: `sender '${message.senderId}' has no live Teams bot`,
    };
  }

  const app = deps.apps.getByAppId(botIdentity.externalId);
  if (!app) {
    return {
      status: 'not-reachable',
      reason: `Teams app '${botIdentity.externalId}' is not running`,
    };
  }

  const appId = botIdentity.externalId;
  const recipient = deps.collective.get(message.recipientId);
  let record: ConversationRecord | undefined;
  let recordKey: string | undefined;

  const mappedKey = deps.state.legionConversationTargets[message.conversationId];
  if (mappedKey) {
    record = deps.state.conversations[mappedKey];
    if (record) recordKey = mappedKey;
  }

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
      recordKey = key;
    } else {
      const key = `${appId}:${message.recipientId}`;
      record = deps.state.conversations[key];
      recordKey = key;

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
            const reason = describeError(error);
            console.error('Teams conversation creation failed', reason);
            return {
              status: 'failed',
              reason: `Teams conversation creation failed: ${reason}`,
            };
          }
          record = {
            conversationId: conversation.id,
            conversationType: 'personal',
            externalId: userIdentity.externalId,
            tenantId: deps.options.tenantId,
            updatedAt: new Date().toISOString(),
          };
          deps.state.conversations[key] = record;
        } else {
          for (const [candidateKey, candidate] of Object.entries(deps.state.conversations)) {
            if (
              candidateKey.startsWith(`${key}:`) &&
              (candidate.conversationType ?? 'personal') === 'personal' &&
              (!record || candidate.updatedAt > record.updatedAt)
            ) {
              record = candidate;
              recordKey = candidateKey;
            }
          }
        }
      }
    }
  }

  if (!record || !recordKey) {
    return {
      status: 'not-reachable',
      reason: `no Teams conversation for '${message.recipientId}'`,
    };
  }

  for (const chunk of chunkMarkdown(message.content, deps.options.maxMessageChars)) {
    const activity = new MessageActivityInput(chunk);
    if (deps.options.aiLabel) activity.addAiGenerated();
    try {
      await app.send(record.conversationId, activity);
    } catch (error) {
      const reason = describeError(error);
      console.error('Teams message delivery failed', reason);
      return { status: 'failed', reason: `Teams message delivery failed: ${reason}` };
    }
  }

  deps.state.legionConversationTargets[message.conversationId] = recordKey;
  if (deps.options.learnOutboundThread) {
    deps.state.legionConversations[recordKey] = message.conversationId;
  }
  await deps.stateStore.save(deps.state);

  return { status: 'delivered' };
}
```

Add the `DeliveryOutcome` type import to the top of `src/outbound.ts`.

- [ ] **Step 4: Update the connector wrapper**

In `src/connector.ts`, import `DeliveryOutcome` from core:

```ts
import type {
  Connector,
  ConnectorContext,
  ConnectorRuntimeDeps,
  DeliveryOutcome,
} from '@legion-collective/core';
```

Replace `deliver` with:

```ts
  async deliver(message: {
    id: string;
    conversationId: string;
    senderId: string;
    recipientId: string;
    replyTo?: string;
    content: string;
    timestamp: string;
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
```

- [ ] **Step 5: Assert the connector outcome**

In `src/connector.test.ts`, in the "starts the pool and delivers outbound messages" test, change the `c.deliver` call to:

```ts
    await expect(c.deliver({
      id: 'm1', conversationId: 'conv-1', senderId: 'assistant',
      recipientId: 'alice', content: 'hello', timestamp: 'now',
    })).resolves.toEqual({ status: 'delivered' });
```

- [ ] **Step 6: Run the tests to verify they pass**

```bash
npx vitest run src/outbound.test.ts src/connector.test.ts
```

Expected: PASS.

- [ ] **Step 7: Commit**

```bash
git add src/outbound.ts src/connector.ts src/outbound.test.ts src/connector.test.ts
git commit -m "feat(teams): report delivery outcomes and learn threads"
```

---

### Task 7: Adopt the delivery contract in the Telegram connector

**Files:**
- Modify: `repos/legion-connector-telegram-v2/src/outbound.ts`
- Modify: `repos/legion-connector-telegram-v2/src/connector.ts`
- Test: `repos/legion-connector-telegram-v2/src/outbound.test.ts`
- Test: `repos/legion-connector-telegram-v2/src/connector.test.ts`

**Interfaces:**
- Consumes: `DeliveryOutcome` from Task 1.
- Produces:
  - `createOutboundHandler(deps): (message) => Promise<DeliveryOutcome>`
  - `TelegramConnector.canDeliver(participantId: string): boolean` (true for `telegram` identity)
  - `TelegramConnector.deliver(message): Promise<DeliveryOutcome>`

- [ ] **Step 1: Point Telegram at the thread's updated core**

Telegram is a trunk-workflow repo and resolves core from an old worktree. Repoint its untracked dependency symlinks to the thread Legion worktree:

```bash
cd /home/chris/workspace/repos/legion-connector-telegram-v2
ln -sfn /home/chris/workspace/threads/2026-11-22-outbound-communication/legion/packages/core \
  node_modules/@legion-collective/core
ln -sfn /home/chris/workspace/threads/2026-11-22-outbound-communication/legion/packages/types \
  node_modules/@legion-collective/types
```

- [ ] **Step 2: Update the outbound tests**

In `src/outbound.test.ts`, replace the assertions in the `describe('outbound handler')` block:

```ts
  it('sends via the sender’s own bot and learns the thread', async () => {
    const { pairs, handler } = await setup(senderWithBot);
    await expect(handler(msg())).resolves.toEqual({ status: 'delivered' });
    expect(pairs.conversationFor('my_bot', '777')).toBe('conv-9');
    expect(pairs.openChats('my_bot', '777')).toBe(true);
  });

  it('reports not-reachable when the recipient has no telegram identity', async () => {
    const { handler } = await setup(senderWithBot);
    await expect(handler(msg({ recipientId: 'other_service' }))).resolves.toEqual({
      status: 'not-reachable',
      reason: "recipient 'other_service' has no telegram identity",
    });
  });

  it('reports not-reachable when the sender has no bot identity', async () => {
    const { pairs, handler } = await setup(senderNoBot);
    await expect(handler(msg())).resolves.toEqual({
      status: 'not-reachable',
      reason: "sender 'quiet' has no telegram-bot identity (sender-only rule)",
    });
    expect(pairs.conversationFor('my_bot', '777')).toBeUndefined();
  });

  it('reports not-reachable when the sender’s bot is not live', async () => {
    const { pairs, handler } = await setup(senderWithBot);
    await expect(handler(msg({ senderId: 'ghost' }))).resolves.toEqual({
      status: 'not-reachable',
      reason: "bot '@my_bot' is not live in the pool",
    });
    expect(pairs.conversationFor('my_bot', '777')).toBeUndefined();
  });

  it('warns with the chat-open hint and still delivers', async () => {
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    const { handler } = await setup(senderWithBot);
    await expect(handler(msg())).resolves.toEqual({ status: 'delivered' });
    expect(warn).toHaveBeenCalledWith(expect.stringContaining('must message @my_bot first'));
    warn.mockRestore();
  });

  it('reports failed when the platform send throws', async () => {
    const { pool, handler } = await setup(senderWithBot);
    const bot = pool.get('my_bot')!;
    vi.spyOn(bot.bot.api, 'sendMessage').mockRejectedValueOnce(new Error('network down'));
    const log = vi.spyOn(console, 'error').mockImplementation(() => undefined);
    try {
      await expect(handler(msg())).resolves.toEqual({
        status: 'failed',
        reason: 'Error: network down',
      });
    } finally {
      log.mockRestore();
    }
  });
```

- [ ] **Step 3: Run the outbound tests to verify they fail**

```bash
cd /home/chris/workspace/repos/legion-connector-telegram-v2
npx vitest run src/outbound.test.ts
```

Expected: FAIL because the handler resolves to `undefined`.

- [ ] **Step 4: Implement the Telegram outcomes**

In `src/outbound.ts`, add:

```ts
import type { DeliveryOutcome } from '@legion-collective/core';
```

Change the handler signature:

```ts
export function createOutboundHandler(
  deps: OutboundDeps,
): (message: OutboundMessage) => Promise<DeliveryOutcome> {
```

Replace the early returns and the send/learning block:

```ts
    if (!chatId) {
      console.debug(
        `[telegram-v2] outbound dropped — recipient '${message.recipientId}' has no telegram identity`,
      );
      return {
        status: 'not-reachable',
        reason: `recipient '${message.recipientId}' has no telegram identity`,
      };
    }

    if (!username) {
      console.warn(
        `[telegram-v2] outbound dropped — sender '${message.senderId}' has no telegram-bot identity (sender-only rule)`,
      );
      return {
        status: 'not-reachable',
        reason: `sender '${message.senderId}' has no telegram-bot identity (sender-only rule)`,
      };
    }

    const poolBot = deps.pool.get(username);
    if (!poolBot) {
      console.warn(`[telegram-v2] outbound dropped — bot '@${username}' is not live in the pool`);
      return {
        status: 'not-reachable',
        reason: `bot '@${username}' is not live in the pool`,
      };
    }

    // The chat-open precheck remains a warning, not a hard failure.
    if (!deps.pairs.openChats(username, chatId)) {
      console.warn(
        `[telegram-v2] user must message @${username} first — chat ${chatId} has not opened a session with this bot`,
      );
    }

    try {
      await sendMarkdown(poolBot.bot.api, chatId, message.content);
    } catch (err) {
      console.error(`[telegram-v2] outbound send to chat ${chatId} failed:`, err);
      return {
        status: 'failed',
        reason: err instanceof Error ? err.message : String(err),
      };
    }

    deps.pairs.remember(username, chatId, message.conversationId);
    return { status: 'delivered' };
```

- [ ] **Step 5: Update the connector wrapper and add `canDeliver`**

In `src/connector.ts`, add `DeliveryOutcome` to the core type import:

```ts
import type {
  Connector,
  ConnectorContext,
  ConnectorRuntimeDeps,
  DeliveryOutcome,
} from '@legion-collective/core';
```

Add `canDeliver` before `deliver`:

```ts
  canDeliver(participantId: string): boolean {
    return Boolean(
      this.deps.collective.get(participantId)?.identities?.some(
        (identity) => identity.connector === 'telegram',
      ),
    );
  }
```

Replace `deliver` with:

```ts
  async deliver(message: {
    id: string;
    conversationId: string;
    senderId: string;
    recipientId: string;
    content: string;
    timestamp: string;
  }): Promise<DeliveryOutcome> {
    if (!this.pool) {
      return { status: 'not-reachable', reason: 'telegram connector is not started' };
    }
    try {
      const outbound = createOutboundHandler({
        collective: this.deps.collective as unknown as Parameters<
          typeof createOutboundHandler
        >[0]['collective'],
        pool: this.pool,
        pairs: this.pairs,
      });
      return await outbound(message);
    } catch (err) {
      console.error('[telegram-v2] deliver failed:', err);
      return {
        status: 'failed',
        reason: err instanceof Error ? err.message : String(err),
      };
    }
  }
```

- [ ] **Step 6: Add the connector `canDeliver` test**

In `src/connector.test.ts`, add:

```ts
  it('reports deliverability from telegram identities', () => {
    const { deps } = makeDeps({
      participants: [
        { id: 'chris', identities: [{ connector: 'telegram', externalId: '777' }] },
        { id: 'worker', identities: [{ connector: 'telegram-bot', externalId: 'my_bot' }] },
      ],
    });
    const c = connector({}, deps);
    expect(c.canDeliver?.('chris')).toBe(true);
    expect(c.canDeliver?.('worker')).toBe(false);
    expect(c.canDeliver?.('nobody')).toBe(false);
  });
```

- [ ] **Step 7: Run the tests to verify they pass**

```bash
npx vitest run src/outbound.test.ts src/connector.test.ts
```

Expected: PASS.

- [ ] **Step 8: Commit**

```bash
git add src/outbound.ts src/connector.ts src/outbound.test.ts src/connector.test.ts
git commit -m "feat(telegram): report delivery outcomes"
```

---

### Task 8: Document the contract and finish both repositories

**Files:**
- Modify: `repos/legion-connector-teams/README.md`
- Modify: `repos/legion-connector-teams/docs/deploy.md`
- Modify: `repos/legion-connector-telegram-v2/README.md`
- Modify: `packages/core/src/connectors/Connector.ts` (JSDoc already added in Task 1)

**Interfaces:**
- Consumes: all prior tasks.
- Produces: user-facing docs and merged/pushed branches.

- [ ] **Step 1: Document Teams outbound initiation**

In `README.md`, add a row to the options table after `maxMessageChars`:

```md
| `learnOutboundThread` | `true` | Record an outbound-initiated Legion conversation as the chat's current thread, so the Teams reply continues it. |
```

Add an `## Outbound initiation` section after the identity/credentials section:

```md
## Outbound initiation

A participant that owns a `teams-bot` identity can start a conversation with an
existing Teams participant using `communicate({ to, message })`. The target must
already exist in the collective with a `teams` (personal) or `teams-group`
(group chat) identity. The connector reports `delivered`, `accepted`,
`not-reachable`, or `failed`; a failed send makes the synchronous `communicate`
call return an error with a sanitized reason.

On a successful send the connector records the Legion conversation as the chat's
current thread by default (`learnOutboundThread: true`), so the Teams user's next
reply continues that conversation. Set `learnOutboundThread: false` to keep the
previous thread; the conversation target map is still updated. Cold outreach,
rosters, retries, and delivery receipts are not supported yet.
```

Update the existing "Proactive delivery is best effort" paragraph to mention that failures now surface to the caller and that thread learning is on by default.

- [ ] **Step 2: Document the troubleshooting change**

In `docs/deploy.md`, replace the 403 bullet with:

```md
- Check for 403 delivery or conversation-creation errors, including an app not installed for the user. Install that bot manually and retry; outbound failures are logged as `Teams message delivery failed` or `Teams conversation creation failed` and returned to the initiating `communicate` call as a `failed` or `not-reachable` reason.
```

- [ ] **Step 3: Document the Telegram outcome contract**

In `repos/legion-connector-telegram-v2/README.md`, update the "Bot-agnostic delivery" bullet:

```md
- **Bot-agnostic delivery with outcomes.** Any pool bot can deliver to any
  registered user's chat. The first message must come from the user (Telegram
  only lets a bot message users who have started a chat with it). If the user
  has never messaged this bot, the outbound attempt is logged as a warning and
  still proceeds. `deliver` reports `delivered`, `not-reachable` (missing
  identity or live bot), or `failed` (platform send error) so `communicate` can
  surface a real failure.
```

- [ ] **Step 4: Run the full verification in both repos**

```bash
cd threads/2026-11-22-outbound-communication/legion
npm test
npm run typecheck
npm run build
cd ../legion-connector-teams
npm test
npm run typecheck
npm run build
cd /home/chris/workspace/repos/legion-connector-telegram-v2
npm test
npm run typecheck
```

Expected:
- Legion: `1345+ passed`, `21 skipped`, typecheck/build clean.
- Teams: `204+ passed`, `1 skipped`, typecheck/build clean.
- Telegram: full suite green, typecheck clean.

- [ ] **Step 5: Commit docs**

```bash
cd /home/chris/workspace/repos/legion-connector-teams
git add README.md docs/deploy.md
git commit -m "docs: document outbound participant communication"

cd /home/chris/workspace/repos/legion-connector-telegram-v2
git add README.md
git commit -m "docs: document delivery outcomes"
```

- [ ] **Step 6: Merge and push Legion**

```bash
cd /home/chris/workspace/repos/legion
git fetch github --quiet
git merge --ff-only feat/outbound-communication
git push github main
```

Expected: `main` fast-forwards to the feature branch and pushes to GitHub.

- [ ] **Step 7: Rebuild `repos/legion` and repoint Telegram**

```bash
cd /home/chris/workspace/repos/legion
npm run build
cd /home/chris/workspace/repos/legion-connector-telegram-v2
ln -sfn /home/chris/workspace/repos/legion/packages/core node_modules/@legion-collective/core
ln -sfn /home/chris/workspace/repos/legion/packages/types node_modules/@legion-collective/types
npm run typecheck
```

- [ ] **Step 8: Merge and push Teams**

```bash
cd /home/chris/workspace/repos/legion-connector-teams
git fetch github --quiet
git merge --ff-only feat/outbound-communication
git push github main
npm run build
```

Expected: `main` fast-forwards, pushes to GitHub, and `dist/` is rebuilt for the live connector.

- [ ] **Step 9: Clean up the thread worktrees**

```bash
cd /home/chris/workspace
git -C repos/legion worktree remove threads/2026-11-22-outbound-communication/legion
git -C repos/legion-connector-teams worktree remove threads/2026-11-22-outbound-communication/legion-connector-teams
git -C repos/legion branch -d feat/outbound-communication
git -C repos/legion-connector-teams branch -d feat/outbound-communication
```

Only remove the worktrees after both merges are confirmed and the branches are ancestors of `main`.

---

### Task 9: Grant the live assistant outbound tools and validate end to end

**Files:**
- Modify: `/tmp/opencode/teams-socket-live/.legion/collective/participants/assistant.json`
- No tracked files.

**Interfaces:**
- Consumes: merged Legion core, merged Teams connector, live workspace config.
- Produces: a live assistant that can `communicate` and `list_participants`, plus recorded validation evidence.

- [ ] **Step 1: Stop the live Legion process**

Find and stop the running process. Do not edit participant files while Legion is running; the in-memory collective will overwrite the file.

```bash
pgrep -af 'packages/runtime/bin/legion.js'
kill <pid>
```

- [ ] **Step 2: Grant the assistant tools**

Edit `/tmp/opencode/teams-socket-live/.legion/collective/participants/assistant.json` so its `tools` object is:

```json
{
  "file_list": "auto",
  "file_read": "auto",
  "file_write": "requires_approval",
  "communicate": "auto",
  "list_participants": "auto"
}
```

Keep the existing `identities`, `model`, `systemPrompt`, and `maxIterations` fields.

- [ ] **Step 3: Start the merged Legion**

```bash
cd /home/chris/workspace/repos/legion
LEGION_WORKSPACE=/tmp/opencode/teams-socket-live \
LEGION_SECRETS_KEY='local-teams-socket-secrets-key-0123456789' \
node packages/runtime/bin/legion.js
```

Expected: web UI at `http://127.0.0.1:3013`; logs show `[legion] external connector 'teams' started`.

- [ ] **Step 4: Validate a personal outbound message**

In the web UI, log in as `operator` and message `assistant`:

> Use `list_participants` to find `teams-chris-willingham-1j2gcm`, then use `communicate` to send: "Outbound personal test from Legion".

Expected: the Teams personal chat receives the message. The assistant's `communicate` result is `success`. Reply from Teams with "Personal reply test"; the reply appears in the **same** Legion conversation in the web UI, not a previous one.

- [ ] **Step 5: Validate a group outbound message**

Message `assistant`:

> Use `communicate` to send "Outbound group test from Legion" to `teams-group-group-pxwcwu`.

Expected: the Teams group receives the message. A group reply appears in the same initiated Legion conversation.

- [ ] **Step 6: Validate a visible delivery failure**

As `operator` (who owns no Teams bot), use `communicate` directly to `teams-chris-willingham-1j2gcm`.

Expected: the tool returns `status: 'error'` with a reason like `sender 'operator' has no live Teams bot`. No Teams message is sent.

- [ ] **Step 7: Record the evidence**

Record the timestamp, the three outcomes, and whether reply continuity held. If any step fails, keep the live process and logs for debugging instead of claiming completion.

---

## Self-Review

- **Spec coverage:** Core contract (§4), Teams (§5), Web/Telegram (§6), policy (§7), data flows (§8), error handling (§9), tests (§10), compatibility (§11) are all mapped to Tasks 1–9. Out-of-scope items (§12) are not implemented.
- **Placeholder scan:** No TBD/TODO placeholders; every code step includes the exact code.
- **Type consistency:** `DeliveryOutcome`, `DeliveryStatus`, `canDeliver`, `getDeliverableConnectors`, `learnOutboundThread`, and `deliverMessage` use the same names and signatures across tasks. `WebConnector`, `TeamsConnector`, and `TelegramConnector` all return `Promise<DeliveryOutcome>`.
