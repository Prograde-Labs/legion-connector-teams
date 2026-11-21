# Legion Teams Connector — Design

**Date:** 2026-11-21
**Status:** Approved for planning
**Package:** `@legion-collective/connector-teams`
**Repo:** `legion-connector-teams`

## 1. Context

Legion v2 hosts agents, services, and human users in one process. External channels
are **connectors**: boundary modules that translate a platform's traffic into Legion
tool calls and Legion messages into platform output. The core knows nothing about
external identity providers.

Today the only external connector is Telegram v2
(`@legion-collective/connector-telegram`), which is the architectural template:

- named `connector(options, deps)` factory, loaded by `loadExternalConnectors`
- `start(ctx)`, `deliver(message)`, `stop()`
- inbound: resolve a platform identity to a Legion participant, then call
  `ctx.streamTool(participantId, 'communicate', …)`
- outbound: `deliver()` reverse-maps the Legion participant to a platform chat and
  sends through the sender's platform identity
- approvals: render cards from the `communicate` result's `approvalRequests`
- connector-owned state for conversation continuity and pairing

This spec adds Microsoft Teams as a second external connector, with **one Teams bot
per participant** (Telegram v2 parity).

### Goals

- Personal 1:1 chats between a Teams user and a Legion participant.
- One bot per participant: the user sees distinct agents in Teams.
- Plain text in and out, `/new`, configurable streaming detail, and approval cards.
- Configurable unknown-user handling: pairing code, auto-provision, reject, or fallback.
- Both HTTP (`/api/messages`) and Socket Mode transports.
- Connector-owned persistent state for conversation continuity across restarts.
- A provisioning script that registers a bot per participant.

### Non-goals (first version)

- Group chats and team channels. A follow-up adds group chats as a participant
  (the group `conversationId` becomes a Legion user participant).
- File attachments.
- OAuth/SSO and delegated Microsoft Graph access.
- Multi-tenant distribution.
- Tabs, message extensions, meeting events, and Copilot surfaces.
- Graph-only messaging.

## 2. Decisions

| Decision | Choice | Rationale |
| --- | --- | --- |
| SDK | Microsoft Teams SDK (`@microsoft/teams.apps` 2.2.x) | Current Teams-first SDK; `/api/messages`, JWT validation, typing, streaming, Adaptive Cards, proactive `app.send` |
| Topology | One bot per participant | Telegram v2 parity; distinct agents in Teams; matches participant-owned identity model |
| Transport | HTTP default, Socket Mode opt-in | HTTP is production-stable; Socket Mode avoids a public URL for dev (experimental, new bots only) |
| Bot registration | Teams-managed bot via `teams app create` | No Azure subscription needed; Azure Bot only for OAuth/SSO/Graph |
| Identity | Entra `aadObjectId` | Stable across bots and tenants; `from.id` is per-bot |
| Unknown users | Configurable: `pairing` (default), `auto-provision`, `reject`, `fallback` | Different deployments need different trust models |
| Credentials | Legion SecretStore, key `teams-bot/<appId>` | Same pattern as Telegram v2; no secrets in participant files or git |
| State | Connector-owned JSON under `stateDir` | Fixes Telegram v2's restart amnesia; no core changes |
| Approvals | Adaptive Cards with `Action.Execute` | Native Teams actions; routed to `card.action.*` handlers |
| Pairing | Core connector-agnostic `pair` tool, AuthEngine policy | Any participant/agent/service granted `pair` can pair; no operator-only gate |
| Tool display | `off` / `status` / `cards` | Configurable use of native Teams streaming and cards |

## 3. Architecture

```
Teams client
   │  (HTTPS POST activity, or Socket Mode WebSocket)
   ▼
Teams SDK App (one per bot)
   │  message / install.add / card.action.*
   ▼
TeamsConnector
   ├── Directory       collective.list() → bots (teams-bot identities + secrets)
   ├── BotPool         start/stop/reconcile App instances
   ├── Identity        resolve aadObjectId → Legion user (policies)
   ├── Pairing         codes + pending state; calls the core `pair` tool
   ├── Inbound         commands, typing, communicate
   ├── Communicate     streamTool chunk mapping
   ├── Outbound        deliver() → app.send
   ├── Approvals       card render + action → approval_response
   ├── Render          markdown, chunking, cards
   └── State           conversations, pending codes
   ▼
Legion core: ConnectorContext.callTool / streamTool / registry; `pair` tool
```

The connector is loaded in-process by Legion. `start(ctx)` builds the directory and
starts the pool. `deliver()` is called by `UserDeliveryRuntime` for messages whose
recipient is a user. `stop()` stops Apps and the HTTP server.

Because the loader calls `start()` before registering the connector, any
`registry.setActive()` performed in `start()` is lost. The connector re-activates on
every inbound message, exactly like Telegram v2.

## 4. Package and module layout

```
src/
  index.ts          named factory: connector(options, deps)
  connector.ts      TeamsConnector implements Connector
  options.ts        TeamsConnectorOptions + defaults
  directory.ts      scan collective → bots
  bot-pool.ts       App lifecycle, reconcile
  transport.ts      HTTP server or Socket Mode
  inbound.ts        message / install.add / conversationUpdate handlers
  identity.ts       aadObjectId → participant resolution
  pairing.ts        code generation, TTL, pending codes
  auto-provision.ts create_user + pair via callTool
  commands.ts       /new, /pair, /whoami, /help
  communicate.ts    streamTool consumption and chunk mapping
  outbound.ts       deliver()
  approvals.ts      card render + action handlers
  render.ts         markdown, chunking, card JSON
  state.ts          atomic JSON state file
  secrets.ts        bot credential resolution
  test-helpers.ts   fakes
```

Runtime dependencies: `@microsoft/teams.apps`, `@microsoft/teams.api`,
`@microsoft/teams.cards`, `express`. Legion packages are type-only devDependencies
(same as Telegram v2). Node 22 (Legion's Docker image is `node:22-slim`).

### 4.1 Core change: connector-agnostic `pair` tool

Legion core gains one small tool, usable by every connector:

```ts
{
  name: 'pair',
  description: 'Merge an external identity into a participant.',
  parameters: {
    type: 'object',
    properties: {
      participantId: { type: 'string' },
      connector:     { type: 'string' },   // e.g. 'teams', 'telegram'
      externalId:    { type: 'string' },   // platform identity
    },
    required: ['participantId', 'connector', 'externalId'],
  },
}
```

Behavior:

- require a non-empty `connector` and `externalId`
- require the target participant to exist
- merge the identity into the participant's `identities` array, deduping on
  `(connector, externalId)`; never replace unrelated identities
- persist via `collective.update(participantId, { identities })`
- return `{ status: 'success', data: { participantId, connector, externalId } }`

Authorization is entirely the AuthEngine's: the caller must have `pair` in its tools
map (`auto`, `requires_approval`, or scoped). Add `pair` to the default operator's
management tools so operators can pair out of the box. The tool knows nothing about
Teams; the connector validates its own pairing code before calling it.

This is a Legion core change plus tests; the connector depends on it. The
implementation plan covers both repositories.

## 5. Configuration

```json
{
  "name": "teams",
  "module": "/app/legion-connector-teams/dist/index.js",
  "enabled": true,
  "options": {
    "transport": "http",
    "port": 3978,
    "publicBaseUrl": "https://teams.prograde-labs.com",
    "tenantId": "<tenant-id>",
    "stateDir": "/workspace/.legion/connectors/teams",
    "unknownUserPolicy": "pairing",
    "maxDetail": "standard",
    "toolDisplay": "status",
    "aiLabel": true
  }
}
```

```ts
export interface TeamsConnectorOptions {
  transport?: 'http' | 'socket';                 // default 'http'
  port?: number;                                  // default 3978
  publicBaseUrl?: string;                         // provisioning/logs only
  tenantId?: string;                              // default tenant for all bots
  stateDir?: string;                              // default $LEGION_WORKSPACE/.legion/connectors/teams
  rescanIntervalMs?: number;                      // default 60000
  socketGeos?: string[];                          // default ['amer','emea','apac']
  unknownUserPolicy?: 'pairing' | 'reject' | 'auto-provision' | 'fallback';
  fallbackUserId?: string;                        // required for 'fallback'
  provisionerId?: string;                         // default 'operator'; needs create_user
  pairingParticipantId?: string;                  // default: the sender; needs pair
  pairingTtlSeconds?: number;                     // default 600
  maxDetail?: 'minimal' | 'standard' | 'full';    // default 'standard'
  toolDisplay?: 'off' | 'status' | 'cards';       // default 'status'
  aiLabel?: boolean;                              // default true
  feedback?: boolean;                             // default false; reserved for a later pass
  typingIntervalMs?: number;                      // default 4000; 0 disables
  maxMessageChars?: number;                       // default 20000
}
```

`start()` throws if `deps.secrets` is absent (`LEGION_SECRETS_KEY` required).

## 6. Identity and credentials

### 6.1 Bot identity

A participant that should be reachable in Teams carries:

```json
{ "connector": "teams-bot", "externalId": "<appId>" }
```

Its credentials live in the SecretStore on that participant:

- key: `teams-bot/<appId>`
- value: JSON `{ "appId": "…", "clientSecret": "…", "tenantId": "…" }`

The bot display name is the participant's `name`. The bot's Teams app manifest
uses the same name.

### 6.2 User identity

A human user participant carries:

```json
{ "connector": "teams", "externalId": "<aadObjectId>" }
```

`aadObjectId` is the Entra object ID. It is stable across bots, whereas
`activity.from.id` is scoped to one bot. If `aadObjectId` is absent (personal
accounts), fall back to `from.id` and record which one was used.

### 6.3 Directory

`Directory.rebuild()` scans `collective.list()`:

- keep participants with `status !== 'retired'`
- for each `teams-bot` identity, resolve `teams-bot/<appId>` from the SecretStore
- validate the JSON; a missing or malformed secret logs a redacted warning and skips
  that bot
- duplicate `appId` across participants: first wins, warn

Rebuild on `start()` and every `rescanIntervalMs`. A rebuild that changes the bot set
triggers `BotPool.reconcile()`.

### 6.4 Outbound routing

`deliver()` receives the recipient (a user) and the sender. The bot is the
**sender's** `teams-bot` identity. The conversation record is keyed
`(appId, participantId)`, because each bot has its own conversation ID with the user.

## 7. Transport and bot pool

### 7.1 HTTP (default)

One Express app and one `http.Server` on `options.port`, bound to `0.0.0.0`.

Each bot gets its own `App`:

```ts
new App({
  clientId: appId,
  clientSecret,
  tenantId,
  httpServerAdapter: new ExpressAdapter(expressApp),
  messagingEndpoint: `/teams/${appId}/api/messages`,
  serviceUrl: 'https://smba.trafficmanager.net/teams',
});
await app.initialize();
```

The Express server listens once. Each bot registration's messaging endpoint is
`<publicBaseUrl>/teams/<appId>/api/messages`. `GET /health` returns 200 without auth.
Proactive sends use the App's configured `serviceUrl` (public cloud default
`https://smba.trafficmanager.net/teams`); sovereign clouds set `cloud`/`serviceUrl`.

For HTTP, a directory change rebuilds the server: stop, recreate the Express app,
initialize all Apps, listen. Rebuilds are rare and brief. Socket Mode reconciles
Apps individually instead (section 7.3).

### 7.2 Socket Mode (opt-in)

```ts
new App({ clientId, clientSecret, tenantId, socketMode: { geos: socketGeos } });
await app.start();
```

No HTTP listener, no tunnel. Each bot opens one socket per geo (default three).
Socket Mode is **experimental** and "new bots only"; it is recommended for
development. `transport: 'socket'` ignores `port` and `publicBaseUrl`.

### 7.3 BotPool

- `start()`: create Apps for every directory bot
- `reconcile(next)`: add new bots, stop removed bots, restart changed credentials
- `stop()`: stop every App and the HTTP server
- per-bot failures are logged and isolated; one bad token does not take down the pool
- secrets are redacted from all log/error text

## 8. Inbound flow

Per-App handlers are registered with a closure over the bot:

```ts
app.on('message', ctx => onMessage(bot, ctx));
app.on('install.add', ctx => onInstall(bot, ctx));
app.on('conversationUpdate', ctx => onConversationUpdate(bot, ctx));
app.on('card.action.approve', ctx => onApprovalAction(bot, ctx, 'approve'));
app.on('card.action.reject',  ctx => onApprovalAction(bot, ctx, 'reject'));
```

### 8.1 Resolution

For `activity.from.aadObjectId` (or `from.id` fallback):

1. `collective.findByIdentity('teams', aadObjectId)` — covers explicit identities and
   anyone paired through the core `pair` tool
2. `fallbackUserId` when `unknownUserPolicy === 'fallback'`
3. policy:
   - `pairing` (default): reply with a code and the user's Teams ID
   - `auto-provision`: `create_user` then `pair` (section 8.5)
   - `reject`: reply with the Teams ID and a registration hint

A retired participant is treated as unknown. Resolution is fail-closed: no tool call
happens until a participant is resolved.

### 8.2 After resolution

- upsert the conversation record `(appId, participantId)` from the activity:
  `conversationId`, `tenantId` (`channelData.tenant.id`), `serviceUrl`,
  `aadObjectId`, `displayName`
- `ctx.registry.setActive(participantId, 'teams')`
- plain text → `communicate` (section 9)
- commands → handled locally (section 8.3)

### 8.3 Commands

- `/new` — clear `legionConversations[appId:participantId]`; reply "Started a new conversation."
- `/pair <code> <participantId>` — validate the pending code (TTL, single use), then
  `ctx.callTool(pairingParticipantId ?? senderId, 'pair', { participantId, connector: 'teams', externalId: aadObjectId })`.
  Delete the pending code on success; reply with confirmation. Surface auth errors
  and `pending_approval` results to the user.
- `/whoami` — reply with participant ID, Teams ID, bot appId, Legion conversation ID.
- `/help` — list commands.

Pairing is authorized by the AuthEngine through the core `pair` tool: any participant
granted `pair` can do it. By default the connector calls the tool as the sender (the
human who typed `/pair`). Set `pairingParticipantId` to route pairing through an agent
or service instead; that participant's `pair` policy is then the one enforced.

### 8.4 Pairing codes

- 6 characters, Crockford base32, ambiguous characters removed
- random, single-use, TTL `pairingTtlSeconds` (default 600)
- bound to `(appId, aadObjectId)`
- reuse an unexpired code for the same pair instead of issuing a new one
- stored in `pendingPairs`; expired entries are pruned on read/write

The unknown-user reply:

> Not registered with Legion yet. Ask someone with the `pair` tool to run
> `/pair <code> <participant-id>`. Your Teams ID: `<aadObjectId>`.

### 8.5 Auto-provision

Two tool calls as the configured participants:

1. `ctx.callTool(provisionerId, 'create_user', { id, name })` where
   `id = teams-<slug(displayName)>-<hash6(aadObjectId)>` and
   `name = activity.from.name`
2. `ctx.callTool(pairingParticipantId ?? provisionerId, 'pair', { participantId: id, connector: 'teams', externalId: aadObjectId })`

`create_user` does not accept identities; `pair` adds the Teams identity, so
subsequent resolution works through `collective.findByIdentity`. If either call returns
`pending_approval`, tell the user approval is pending. If either errors, reply with the
error and leave the user unresolved. The provisioner needs `create_user` as `auto`; the
pairing participant needs a `pair` policy that allows the call.

## 9. Communicate and streaming

```ts
const { gen, conversationId } = await ctx.streamTool(userId, 'communicate', {
  to: botOwnerId,
  message: text,
  detail,                 // clamped to options.maxDetail
  conversationId: state.legionConversations[`${appId}:${userId}`],
});
```

Store `conversationId` immediately. Consume `StreamChunk`s:

| Chunk | `minimal` | `standard` | `full` |
| --- | --- | --- | --- |
| `iteration_start` | — | `stream.update('Thinking…')` | same |
| `text_delta` | buffer | `stream.emit(delta)` | `stream.emit(delta)` |
| `reasoning_delta` | — | — | optional status line |
| `message_snapshot` | replace buffer | replace buffer | replace buffer |
| `tool_call_start` | — | `toolDisplay: status` → `stream.update('Calling <tool>…')`; `cards` → Adaptive Card | same |
| `tool_call_args_delta` | — | accumulate for card | same |
| `call_timeout` | notify | notify | notify |
| `stream:done` | send buffered text | `stream.close()` | `stream.close()` |
| `stream:error` | send error | send error | send error |

- Typing: send a `TypingActivityInput` every `typingIntervalMs` while the call runs.
- The two-minute streaming window starts on the **first streamed chunk**, and an
  informative `stream.update()` counts as a chunk. Do slow pre-output work before
  opening the stream, and note that an update emitted after response text has
  accumulated is ignored by Teams.
- `toolDisplay: 'cards'` sends a separate Adaptive Card per `tool_call_start` via
  `send()` rather than `stream.emit()`, so tool cards do not disturb the text stream.
- `stream:done` result handling:
  - `success` — text already streamed; if `minimal`, send the buffer
  - `pending_approval` — render approval cards from `result.approvalRequests`
  - `error` / `rejected` / `timeout` — render the message; `timeout` may include a
    partial response
- Teams streaming is **1:1 only** and has a **two-minute window**. On timeout the SDK
  finalizes the message with the buffered text. Long-run handoff to plain edits is a
  follow-up.
- `aiLabel` adds `addAiGenerated()` to the final activity. `feedback` is off by
  default.

`StreamChunk` has `tool_call_start` and `tool_call_args_delta` but no per-tool result
chunk, so tool cards show the invocation and args, not the tool output.

## 10. Outbound

`deliver(message)`:

1. `sender = collective.get(message.senderId)`; find `teams-bot` identity → `appId`
2. if no bot: warn and drop (the sender is not a Teams-facing participant)
3. find `conversations[appId:message.recipientId]`
4. if missing, try the recipient's `teams` identity and create the personal
   conversation with
   `app.api.conversations.create({ members: [{ id: aadObjectId }], tenantId, channelData: { tenant: { id: tenantId } } })`
   and store the returned conversation ID
5. if still missing: warn and drop
6. render `message.content` as Teams markdown, chunk at `maxMessageChars`
7. `app.send(conversationId, new MessageActivityInput(chunk))`, optionally
   `.addAiGenerated()`
8. on 403 (`MessageWritesBlocked`, not installed), log and continue

Proactive messages before the bot is installed for the user cannot succeed; the
operator must install each bot.

## 11. Approvals

When a `communicate` result is `pending_approval`, render one Adaptive Card per
`approvalRequests[]` entry in the same chat:

```json
{
  "type": "AdaptiveCard",
  "version": "1.4",
  "body": [
    { "type": "TextBlock", "text": "🔐 Approval requested", "weight": "Bolder" },
    { "type": "FactSet", "facts": [
      { "title": "Tool", "value": "<tool>" },
      { "title": "Args", "value": "<truncated args>" },
      { "title": "Requester", "value": "<requester>" }
    ]}
  ],
  "actions": [
    { "type": "Action.Execute", "title": "Approve", "data": { "action": "approve", "approvalId": "<id>" } },
    { "type": "Action.Execute", "title": "Reject",  "data": { "action": "reject",  "approvalId": "<id>" } }
  ]
}
```

Built with `@microsoft/teams.cards` (`ExecuteAction` + `SubmitData`) and sent with
`MessageActivityInput().addCard('adaptive', card)`.

Handlers `app.on('card.action.approve')` / `app.on('card.action.reject')`:

1. read `activity.value.action.data`
2. resolve the clicking user the same way as inbound messages
3. `ctx.callTool(userId, 'approval_response', { decisions: [{ approvalId, decision }] })`
4. return a Teams message response
5. the agent's continuation arrives through `deliver()` and is sent to the chat

Auth is enforced by the AuthEngine. The operator has `approval_response: auto` by
default; auto-provisioned users may need it granted. Auth errors are surfaced to the
user.

## 12. Persistent state

`stateDir` default `$LEGION_WORKSPACE/.legion/connectors/teams`, file `state.json`,
mode 0600, atomic write (temp file + rename), no secrets:

```json
{
  "version": 1,
  "conversations": {
    "<appId>:<participantId>": {
      "conversationId": "…",
      "tenantId": "…",
      "serviceUrl": "…",
      "aadObjectId": "…",
      "displayName": "…",
      "updatedAt": "…"
    }
  },
  "pendingPairs": {
    "<code>": { "appId": "…", "aadObjectId": "…", "expiresAt": "…" }
  },
  "legionConversations": {
    "<appId>:<participantId>": "<legion-conversation-id>"
  }
}
```

Pairings themselves live in participant files (written by the core `pair` tool); the
connector state only holds pending codes and conversation references.

A corrupt file is renamed to `state.json.bak-<timestamp>` and replaced with a fresh
state; the connector logs a warning. State is written after every mutation (small file,
low frequency).

## 13. Provisioning

`scripts/provision-teams-bot.mjs` (Node 22) wraps the Teams Developer CLI:

```
teams app create \
  --name <participant-name> \
  --endpoint <publicBaseUrl>/teams/<appId>/api/messages \
  --env <env-file>
```

Steps:

1. read the participant file; require or create the `teams-bot` identity
2. run `teams app create` (Teams-managed bot; no Azure subscription)
3. write a personal-scope manifest + color/outline icons + zip:
   `bots: [{ botId, scopes: ["personal"], supportsFiles: false }]`
4. store `{ appId, clientSecret, tenantId }` in the SecretStore under
   `teams-bot/<appId>` (via a `FileSecretStore` helper or the `set_secret` API)
5. print the install link (`teams app get <appId> --install-link`)

A `provision-all` wrapper loops over participant IDs. The operator installs each bot
once in Teams. Requires tenant sideloading enabled.

## 14. Testing

Vitest, same patterns as Telegram v2:

- `test-helpers.ts`: fake `App` capturing `send`/`stream`, fake activities,
  `InMemorySecretStore`, `StubCollective`, temp `stateDir`
- directory: token-resolvable bots, retired excluded, duplicate appId first-wins,
  malformed secret skipped
- bot-pool: start/stop/reconcile, bad credential isolation, redaction
- identity: all four policies, retired participant, missing aadObjectId fallback
- pairing: code shape, TTL, single use, reuse, configured pairing participant,
  AuthEngine outcomes (`auto`, `requires_approval`, `hidden`)
- commands: `/new`, `/pair`, `/whoami`, `/help`
- core `pair` tool: merge, dedupe, missing participant, empty ids, persistence
- communicate: detail clamping, minimal/standard/full, snapshot, tool status/cards,
  timeout, error
- outbound: sender bot resolution, missing bot, missing conversation, create
  conversation, chunking, 403 handling
- approvals: card shape, action routing, `approval_response` call, auth error
- state: round-trip, atomic write, corruption recovery

Optional live integration behind `LEGION_TEAMS_INTEGRATION=1` using Socket Mode.

## 15. Deployment

- Build `dist/`; mount read-only into the Legion container at
  `/app/legion-connector-teams`.
- Add `teams` to `.legion/config.json` with the module path and options.
- Expose the connector port (HTTP) through the reverse proxy/tunnel at
  `teams.prograde-labs.com`.
- `LEGION_SECRETS_KEY` must be set; bot secrets live in the SecretStore.
- For Socket Mode, no port or public endpoint is needed.

The free Microsoft 365 Business Basic trial provides a tenant for about one month.
A paid seat is required for durable use.

## 16. Security

- Inbound JWT validation is the SDK's per-App auth (`clientId`/`tenantId`).
- Single-tenant deployment; no multi-tenant tokens.
- `aadObjectId` is the canonical identity; display names are untrusted.
- Pairing codes are random, single-use, and time-limited.
- `/pair` is gated by the AuthEngine's `pair` tool policy; any participant, agent, or
  service granted `pair` can pair, and `hidden`/`requires_approval` behave normally.
- Unknown users are fail-closed unless a policy explicitly allows them.
- Bot secrets are read from the SecretStore and redacted from logs/errors.
- Approval actions are executed as the clicking user, so AuthEngine policy applies.
- State contains IDs only, never secrets.

## 17. Risks and open items

| Risk | Mitigation |
| --- | --- |
| Free trial tenant expires | Document the paid-seat requirement; keep tenant-specific config out of git |
| Teams-managed bots + Socket Mode are experimental | Pin SDK versions; HTTP is the production default; Socket Mode is dev-only |
| One bot per participant means N registrations/installs | Provisioning script + install links; document the per-bot install step |
| Proactive messages fail before install | Log 403s; operator installs each bot |
| Streaming two-minute limit | SDK finalizes with buffered text; long-run edit handoff is a follow-up |
| Teams markdown subset | Render conservatively; chunk long messages |
| `create_user` has no identities | Auto-provision calls `create_user` then the core `pair` tool |
| Core `pair` tool is a Legion change | Implement and test it in Legion first; the connector depends on it |
| Pairing policy misconfigured | Document `pair` grants; `hidden`/`requires_approval` fail visibly |
| Group chats later | Map group `conversationId` to a user participant; sender attribution is a separate design |

## 18. Future work

- Group chats as participants.
- Team channels with @mentions and threaded replies (RSC permissions, channel install).
- File attachments (Teams file consent / Graph upload).
- Long-run streaming handoff to plain message edits.
- Citations, feedback, suggested prompts, and the AI-generated label as configurable
  enhancements.
- Graph-based automation for app installation.
