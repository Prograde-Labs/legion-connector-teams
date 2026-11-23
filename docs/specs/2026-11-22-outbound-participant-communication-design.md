# Legion — Outbound Participant Communication

**Date:** 2026-11-22
**Status:** Approved for planning
**Packages:** `@legion-collective/core`, `@legion-collective/runtime`, `@legion-collective/connector-teams`
**Builds on:** `docs/specs/2026-11-21-teams-connector-design.md`, `docs/specs/2026-11-22-teams-group-chats-design.md`

## 1. Context

A Legion participant can already send a message to another participant with
`communicate({ to, message })`. Omitting `conversationId` creates a fresh Legion
conversation. The Teams connector's `deliver()` can also cold-start a personal
Teams chat (`conversations.create`) and send directly to a known group
conversation.

In practice, though, proactive outbound communication does not work:

- `UserDeliveryRuntime` only calls connectors returned by
  `ConnectorRegistry.getActiveConnectors()`. External connectors call
  `setActive()` only on inbound, so after a restart a proactive send is dropped.
- The Teams and Telegram outbound paths require the **sender** to own the
  platform bot. Only the bot-owning agent can send; the operator or another
  agent is dropped.
- Teams outbound does not update the chat's current Legion conversation, so the
  external user's reply lands in the previous thread rather than the thread the
  agent initiated.
- Delivery failures are log-only. The initiator still sees `success`.
- An agent needs `communicate` and a way to discover participant IDs; the live
  `assistant` has neither.

This spec defines a connector-agnostic delivery seam in Legion core and applies it
to Teams first. It lets a bot-owning participant start a conversation with an
existing external participant and learn the thread so replies continue it.

### Goals

- A bot-owning Legion participant can initiate a message to an existing Teams
  personal user or group participant without waiting for inbound traffic.
- The core no longer relies on inbound-derived "active" state to decide whether a
  connector can deliver.
- `communicate` reports a delivery failure to the initiator instead of silently
  succeeding.
- On a successful outbound send, the connector records the initiated Legion
  conversation as the chat's current thread (configurable, default on), so the
  external user's reply continues it.
- Existing inbound, approvals, typing, streaming, and personal/group behavior
  remain intact.
- Web and Telegram connectors adopt the same contract.

### Non-goals

- Cold outreach to a Teams user or group that has never messaged Legion.
- A managed external-contact roster.
- A persistent outbox, retry queue, or delivery receipts.
- A web UI participant picker for starting external conversations.
- Teams channels, per-human participants in group chats, or multi-bot fan-out.
- Changing the sender-only rule: only a participant that owns a platform bot may
  initiate.

## 2. Decisions

| Decision | Choice | Rationale |
| --- | --- | --- |
| Scope | Generic core seam, Teams first | Fixes the root cause (`active` gating) without building an outbox |
| Deliverability | `Connector.canDeliver(participantId)` predicate | Explicit and state-aware; `setActive` stays presence-only |
| Sender rule | Sender must own the platform bot | Keeps attribution exact; the assistant is the only current bot owner |
| Targets | Existing Legion participants only | No roster or cold-target work this round |
| Delivery outcome | `delivered` / `accepted` / `not-reachable` / `failed` | Truthful result for the initiator without retries |
| Thread learning | Configurable per connector; default on | Matches Telegram's existing behavior; operators can opt out |
| Error surfacing | `MessageRouter` maps a failed `void` to `status: 'error'` | `communicate` already surfaces errors to the caller |
| Fire-and-forget | Still returns `dispatched` | Async dispatch cannot await the connector outcome |
| State | No schema change | Reuses `legionConversations` and `legionConversationTargets` |

## 3. Architecture

```
Legion participant (bot owner)
   │  communicate({ to: <external participant>, message })
   ▼
MessageRouter
   │  creates/reuses a Legion conversation
   ▼
UserDeliveryRuntime (recipient is a user participant)
   │  ConnectorRegistry.getDeliverableConnectors(recipient)
   ├── canDeliver? → Teams / Telegram / Web
   ├── deliver() → DeliveryOutcome
   └── aggregate → RuntimeResult.void.delivery
   ▼
MessageRouter.completeResponse
   │  failed outcome → status: 'error'
   ▼
communicate returns success or a delivery error

TeamsConnector.deliver
   ├── sender's teams-bot identity + owner check
   ├── resolve target (target map → group → personal → create → fallback)
   ├── app.send(...)
   ├── learn thread: target map always; legionConversations if enabled
   └── return delivered / accepted / not-reachable / failed
```

## 4. Core delivery contract (Legion)

### 4.1 `DeliveryOutcome`

Add to `packages/core/src/connectors/Connector.ts`:

```ts
export type DeliveryStatus = 'delivered' | 'accepted' | 'not-reachable' | 'failed';

export interface DeliveryOutcome {
  status: DeliveryStatus;
  /** Human-readable reason for not-reachable or failed. */
  reason?: string;
}
```

- `delivered`: the platform accepted the message.
- `accepted`: Legion accepted responsibility but did not confirm a live platform
  send (used by the web connector when the recipient has no open socket; the
  message is persisted and the UI reloads it).
- `not-reachable`: no connector or bot can reach the recipient.
- `failed`: a connector attempted delivery and the platform call failed.

### 4.2 `Connector` interface

Change `deliver` to return the outcome, and add an optional predicate:

```ts
export interface Connector {
  readonly name: string;

  /**
   * Whether this connector can deliver to the participant. Omit to mean
   * "always ask me"; implemented connectors should base this on the
   * participant's identities and their own liveness.
   */
  canDeliver?(participantId: string): boolean;

  deliver(message: {
    id: string;
    conversationId: string;
    senderId: string;
    recipientId: string;
    replyTo?: string;
    content: string;
    timestamp: string;
  }): Promise<DeliveryOutcome>;

  start(ctx: ConnectorContext): Promise<void>;
  stop(): Promise<void>;
}
```

`canDeliver` is optional for compatibility; the registry treats an absent
predicate as `true`. All in-repo connectors implement it.

### 4.3 `ConnectorRegistry`

Add:

```ts
getDeliverableConnectors(participantId: string): Connector[] {
  return [...this.connectors.values()].filter(
    (connector) => connector.canDeliver?.(participantId) ?? true,
  );
}
```

`getActiveConnectors()` remains for presence (web sockets) and is no longer used
by `UserDeliveryRuntime`. `setActive()` / `clearActive()` semantics are
unchanged.

### 4.4 `UserDeliveryRuntime`

Use the deliverable set, collect outcomes, and return a `delivery` field:

```ts
const connectors = this.connectorRegistry?.getDeliverableConnectors(this.participantId) ?? [];
if (connectors.length === 0) {
  return {
    kind: 'void',
    delivery: {
      status: 'not-reachable',
      reason: `no connector can reach ${this.participantId}`,
    },
  };
}

const outcomes: DeliveryOutcome[] = [];
for (const connector of connectors) {
  try {
    outcomes.push(await connector.deliver({ ... }));
  } catch (error) {
    outcomes.push({
      status: 'failed',
      reason: error instanceof Error ? error.message : String(error),
    });
  }
}

const successful = outcomes.find(
  (outcome) => outcome.status === 'delivered' || outcome.status === 'accepted',
);
if (successful) {
  return { kind: 'void', delivery: successful };
}

const failed = outcomes.find((outcome) => outcome.status === 'failed');
const reasons = outcomes.map((outcome) => outcome.reason ?? outcome.status).join('; ');
return {
  kind: 'void',
  delivery: {
    status: failed ? 'failed' : 'not-reachable',
    reason: reasons,
  },
};
```

If multiple connectors deliver, the first successful outcome is reported; fan-out
is preserved from the current behavior.

### 4.5 `RuntimeResult`

Extend the `void` variant:

```ts
| {
    kind: 'void';
    delivery?: DeliveryOutcome;
    actions?: MiddlewareActionResult[];
  }
```

`delivery` is optional; other runtimes (agent, mock) continue to return a plain
`void`.

### 4.6 `MessageRouter`

In `completeResponse`, replace the `void` success mapping:

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

This covers both the synchronous (`sendStream`/`send`) and asynchronous
(`replyTo`) paths. The async path still returns `dispatched` to the caller before
delivery completes; its outcome is not surfaced.

### 4.7 `communicate`

No signature change. A synchronous call now returns `status: 'error'` with the
connector's reason when delivery fails. A fire-and-forget call remains
`dispatched`.

## 5. Teams connector

### 5.1 `canDeliver`

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

The recipient must be an existing Legion participant with a Teams identity.

### 5.2 `deliver`

`deliverMessage` returns a `DeliveryOutcome`:

1. Resolve the sender's `teams-bot` identity and require
   `ownerOf(appId) === senderId`. Otherwise:
   `{ status: 'not-reachable', reason: 'sender ... has no live Teams bot' }`.
2. Resolve the target:
   - `legionConversationTargets[message.conversationId]` → record.
   - Else a `teams-group` identity → record keyed
     `<appId>:<participantId>:<externalId>`; send directly to `externalId` when
     no record exists.
   - Else a `teams` identity → record keyed `<appId>:<participantId>`; if
     missing, `app.api.conversations.create({ members: [{ id: externalId }],
     tenantId, channelData })`.
   - Else the latest personal record for that participant, if any.
3. If no record or app: `{ status: 'not-reachable', reason }`.
4. Chunk, add the AI label, and `app.send`. On failure:
   `{ status: 'failed', reason }`.
5. On success, learn the thread (5.3) and return `{ status: 'delivered' }`.

A `403` / not-installed response is a `failed` outcome with a reason. The
initiator sees it; nothing is retried.

### 5.3 Thread learning

On a successful send:

- Always set `legionConversationTargets[message.conversationId] = key`, where
  `key` is the platform conversation key resolved above.
- If `options.learnOutboundThread` is `true` (default), also set
  `legionConversations[key] = message.conversationId`.

This makes the initiated Legion conversation the chat's current thread, so the
external user's next inbound message continues it. With
`learnOutboundThread: false`, the chat keeps its previous thread and only the
target map is updated.

State is saved once after the update. No schema change.

### 5.4 Option

Add `learnOutboundThread`:

| Option | Default | Purpose |
| --- | --- | --- |
| `learnOutboundThread` | `true` | Record an outbound-initiated Legion conversation as the chat's current thread |

`resolveOptions` coerces it to a boolean and exposes it on `ResolvedOptions`.

## 6. Other connectors

### 6.1 Web

- `canDeliver`: recipient has a `web` identity.
- `deliver`: `delivered` when a socket exists and the send succeeds;
  `accepted` when no socket exists (the message is persisted and the UI reloads
  it); `failed` when a socket send throws.

### 6.2 Telegram

- `canDeliver`: recipient has a `telegram` identity.
- `deliver`: `delivered` on success; `not-reachable` when the recipient or
  sender identity is missing, or the bot is not live; `failed` when the platform
  send throws.
- Keep the existing `pairs.remember(...)` thread learning on success.

## 7. Policy and deployment

- The `assistant` agent must be granted `communicate` and `list_participants`
  (via `modify_agent` or the participant config) so it can discover existing
  Teams participants and initiate.
- Only participants with a platform bot identity can initiate; the assistant is
  the only current owner of the Teams bot.
- Scope is existing participants: a Teams user/group must already have a Legion
  participant (from inbound auto-provision/pairing).
- No approval is added for outbound sends in this round; the existing tool policy
  on `communicate` applies.

## 8. Data flows

### 8.1 Agent initiates a Teams personal message

1. Assistant calls `communicate({ to: 'teams-chris-...', message })`.
2. `MessageRouter` creates a new Legion conversation.
3. `UserDeliveryRuntime` asks `getDeliverableConnectors('teams-chris-...')`;
   Teams returns true.
4. `TeamsConnector.deliver` resolves the assistant's bot, reuses the personal
   record (or creates the chat), sends, and returns `delivered`.
5. `deliverMessage` sets the target map and, with learning on,
   `legionConversations[key] = new conversation`.
6. The human replies in Teams; inbound looks up `legionConversations[key]` and
   continues the initiated conversation.

### 8.2 Agent initiates a group message

1. Assistant targets `teams-group-group-pxwcwu`.
2. Teams `canDeliver` returns true for the `teams-group` identity.
3. `deliver` sends to the group conversation ID; the group record is created if
   absent.
4. Thread learning updates the group key, so the next group message continues the
   initiated conversation.

### 8.3 Delivery failure

1. The sender has no bot identity, the app is not live, or `app.send` fails.
2. The connector returns `not-reachable` or `failed` with a reason.
3. `UserDeliveryRuntime` returns `void` with the outcome.
4. `MessageRouter` returns `status: 'error'`.
5. `communicate` returns `status: 'error'` with the reason.

## 9. Error handling

- Connector exceptions are caught by `UserDeliveryRuntime` and converted to
  `failed` outcomes; one connector's failure does not prevent another from
  delivering.
- A missing identity or app is `not-reachable`; a platform call failure is
  `failed`.
- Delivery reasons must not include credentials or secrets; reuse
  `describeError` where applicable.
- Fire-and-forget delivery failures remain log-only because the caller has already
  received `dispatched`.

## 10. Testing

### Core

- `ConnectorRegistry.getDeliverableConnectors` includes connectors whose
  predicate returns true and excludes those whose predicate returns false.
- `UserDeliveryRuntime` delivers without `setActive`; fans out to multiple
  connectors; returns `not-reachable` when none can reach; returns `failed` when
  all attempts fail; returns success when any connector delivers or accepts.
- `MessageRouter` maps a failed `void` delivery to `status: 'error'` and leaves
  successful/absent delivery as success.
- `communicate` surfaces the error in a synchronous call and keeps `dispatched`
  for fire-and-forget.

### Teams

- `canDeliver` is true for `teams` and `teams-group` identities and false
  otherwise.
- `deliver` returns `delivered` on a successful send, `not-reachable` for a
  non-bot sender / missing app / missing target, and `failed` for create/send
  errors.
- Thread learning: target map always set; `legionConversations` updated when
  `learnOutboundThread` is true and unchanged when false.
- Personal cold-start (`conversations.create`) and group direct send.
- Existing inbound, approvals, streaming, and group tests stay green.

### Web and Telegram

- `canDeliver` by identity.
- `deliver` outcomes for live socket / no socket (web) and missing bot / send
  failure (Telegram).

### Integration

- Agent-initiated personal and group messages reach Teams.
- The human's reply continues the initiated Legion conversation.
- A delivery failure is visible to the initiating agent.

## 11. Compatibility and rollout

- `Connector.deliver` changes from `Promise<void>` to
  `Promise<DeliveryOutcome>`. All in-repo connectors are updated in the same
  change. External connectors must also return an outcome; document the new
  contract in the connector docs.
- `canDeliver` is optional; connectors without it are treated as reachable. It
  should be a cheap predicate (identity/liveness checks), not an I/O probe;
  `deliver` remains the source of truth for actual reachability.
- `setActive` / `getActiveConnectors` remain for presence; no behavior change
  for web sessions.
- Connector state schema is unchanged; `learnOutboundThread` is an option, not a
  state field.
- The Teams app manifest, tenant, and bot installation requirements are
  unchanged.

## 12. Out of scope

- Cold targets by raw external identity (AAD object ID, group conversation ID,
  Telegram chat ID).
- A managed external-contact roster.
- A persistent outbox, retries, or delivery receipts.
- A web UI affordance for starting a conversation with an external participant.
- Teams channels, per-human participants, multi-bot fan-out, and file
  attachments.
