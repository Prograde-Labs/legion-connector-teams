# Legion Teams Connector — Conversation Commands and Picker

**Date:** 2026-11-25
**Status:** Approved for planning
**Package:** `@legion-collective/connector-teams`
**Builds on:** `docs/specs/2026-11-21-teams-connector-design.md`,
`docs/specs/2026-11-22-teams-group-chats-design.md`,
`docs/specs/2026-11-24-card-tool-buttons-design.md`

## 1. Context

The Teams connector currently exposes `/new`, `/pair`, `/whoami`, and `/help`.
`/new` clears the chat's Legion conversation mapping, so the next message starts a
fresh conversation. There is no way to see the current thread's usage, compact it, or
return to an earlier thread once `/new` has been used.

This spec adds three commands and a card-driven conversation picker:

- `/switch` — list the current chat participant's active Legion conversations and
  switch the chat to one of them.
- `/compact [instruction]` — compact the current conversation through a configured
  summarizer agent.
- `/context` — show the current conversation and its token/cost totals.

The picker uses an Adaptive Card with `Action.Execute` buttons, matching the
approval and tool-button patterns already in the connector. All work runs as the
chat's participant, so Legion's existing grants, scopes, and visibility rules apply.

### Goals

- Add `/switch`, `/compact`, and `/context` alongside the existing commands.
- `/switch` with no argument sends a card listing the chat participant's active
  conversations, with buttons to choose one.
- `/switch <conversation-id>` switches directly.
- The picker pages five conversations at a time, with a `More…` button.
- Switching rebinds the Teams chat to the chosen Legion conversation and updates the
  proactive-delivery target map.
- `/compact` uses a new `compaction` connector option and supports an optional
  instruction override.
- `/context` reports the conversation title, status, message counts, and token/cost
  totals.
- Auto-provisioned personal users and group participants receive the command tools.
- The Teams app manifest advertises the commands.
- No Legion core changes.

### Non-goals

- Team channels (`conversationType === 'channel'`).
- Switching a conversation that belongs to another participant.
- Renaming, archiving, or deleting conversations from Teams.
- Branch switching (`switch_branch`).
- Per-chat conversation aliases or pinned favourites.
- Rich pagination beyond a single `More…` button.
- Scheduling auto-compaction from Teams.
- Changing `/new`, `/pair`, or `/whoami` semantics.
- Changing how unknown slash commands are handled (they remain ordinary messages).

## 2. Decisions

| Decision | Choice | Rationale |
| --- | --- | --- |
| Command set | `/switch`, `/compact [instruction]`, `/context` | Directly covers the requested features; mirrors Telegram-v2 naming |
| Conversation scope | The chat participant's active conversations | Personal chat → the paired user; group chat → the group participant |
| Listing tool | `list_conversations` called as the chat participant with `{participantId, status: 'active'}` | Core already resolves titles for the viewer and excludes sub-threads |
| Current marker | Compare the chat's `legionConversations[key]` with the listed ids | No new state |
| Card layout | Five conversation buttons plus `More…` when needed | Adaptive Cards renders at most six actions per card |
| Page state | `offset` in the card action payload | Stateless; the payload is validated and clamped |
| Switch action | `Action.Execute` with `data.action = 'switch'` | Matches the SDK's `card.action.<action>` routing and existing approval/tool cards |
| Ack | Return a running card immediately, then update in place | Avoids the Teams invoke timeout and prevents double-clicks |
| Update target | `activity.replyToId`; follow-up card when absent | Same pattern as tool buttons |
| Clicker authority | Personal: the chat participant. Group: any member, as long as the group participant resolves | Switching only rebinds the group's thread; it runs no tool as the human |
| Conversation validation | The chosen id must appear in the chat participant's active list | Fail closed against tampered card payloads |
| Queueing | Card switch work runs through the existing per-chat `SerialQueue` | A switch cannot be overwritten by an in-flight turn's final mapping write |
| State | No schema change; reuse `legionConversations` and `legionConversationTargets` | Switching is a rebind, not a new state concept |
| Compaction caller | New `compaction: {agentId, instruction?}` option | `compact_conversation` needs `agentId` for manual compaction |
| Missing compaction config | `/compact` replies with setup guidance and calls no tool | Explicit failure instead of a core-tool error |
| Context output | Title, status, total/active messages, and `query_usage` totals | Chosen balance of thread detail and token usage |
| Usage failure | Degrade to thread details and report usage as unavailable | `/context` remains useful without `query_usage` |
| Tool grants | Extend auto-provisioning; document manual grants for existing users | New users work out of the box; existing users stay fail-closed |
| Manifest | Add `commandLists` for `personal` and `groupChat`; bump version to `1.1.0` | Teams surfaces the commands and accepts an update |
| Typing | Factor the existing typing tick into `startTyping`; stop before the final command reply | Keeps `/compact` responsive without a typing indicator after the reply |

## 3. Architecture

### 3.1 Files

- `src/commands.ts` — new command cases and help text.
- `src/conversations.ts` (new) — listing, binding, card construction, and
  conversation-resolution helpers.
- `src/identity.ts` — `resolveChatParticipant` for card actions (registered or
  fallback, matching inbound key construction).
- `src/inbound.ts` — pass `sendCard`/`stopTyping` to commands; use `startTyping`.
- `src/connector.ts` — `onConversationAction` for `card.action.switch`.
- `src/bot-pool.ts` — register `card.action.switch` and add it to `AppHandlers`.
- `src/options.ts` — `compaction` option and validation.
- `scripts/provision-teams-bot.mjs` — manifest command list and version bump.
- `README.md`, `docs/deploy.md` — commands, option, grants, and reprovisioning.
- Tests listed in §8.

### 3.2 Command surface

`handleCommand` keeps its current structure and adds three cases.

```
/new
/switch
/switch <conversation-id>
/compact [instruction]
/context
/pair <code> <participant-id>
/whoami
/help
```

`/switch`, `/compact`, and `/context` use the `conversationKey` supplied by
inbound. In a group chat that key is the group participant's key; in a personal
chat it is the user's key, including the fallback suffix when applicable.

`CommandDeps` gains:

```ts
sendCard: (card: IAdaptiveCard) => Promise<void>;
stopTyping?: () => void;
```

`sendCard` sends an Adaptive Card through the current Teams context and applies
the `aiLabel` option. `stopTyping` is idempotent and lets `/compact` stop the
indicator before its final message.

The help text becomes:

```
Commands:
/new — start a new conversation
/switch — list and switch to an existing conversation
/switch <conversation-id> — switch directly
/compact [instruction] — summarize/compact this conversation
/context — show usage and context details
/pair <code> <participant-id> — pair a pending Teams identity
/whoami — show your identity
/help — this message
In a group chat, /new, /switch, /compact, and /context act on the group; /pair runs as you.
```

Unknown slash commands continue to return `false` from `handleCommand` and are
forwarded to `communicate`, preserving existing behaviour.

### 3.3 Conversation listing

`src/conversations.ts` exports:

```ts
export interface ConversationEntry {
  id: string;
  title: string;
  messageCount: number;
  updatedAt: string;
  status: 'active' | 'archived';
}

export async function listConversations(
  ctx: ConnectorContext,
  participantId: string,
): Promise<ConversationEntry[]>;
```

`listConversations` calls:

```ts
ctx.callTool(participantId, 'list_conversations', {
  participantId,
  status: 'active',
});
```

It throws a descriptive error when the result is not `success`. On success it maps
`data.conversations ?? []` (core `ConversationMeta`) to `ConversationEntry`:

- `id` — `meta.id`.
- `title` — `meta.resolvedTitle ?? meta.sharedTitle ?? meta.title ?? meta.id`.
- `messageCount` — `meta.messageCount`.
- `updatedAt` — `meta.updatedAt`.
- `status` — `meta.status`.

Entries are sorted by `updatedAt` descending in the connector as well, so the card
is deterministic even if the store's order changes.

`listConversations` is used by both `/switch` and the card handler. Because the
caller is the chat participant, the core tool resolves titles for that participant
and applies that participant's grants.

### 3.4 Binding a conversation

`src/conversations.ts` exports:

```ts
export function bindConversation(
  state: TeamsState,
  key: string,
  conversationId: string,
): void;
```

Behaviour:

- If `state.legionConversations[key] === conversationId`, do nothing.
- Otherwise, read the previous id. Delete
  `state.legionConversationTargets[previous]` only when it still equals `key`.
- Set `state.legionConversations[key] = conversationId`.
- Set `state.legionConversationTargets[conversationId] = key`.

This mirrors the target-map cleanup `/new` already performs. If two Teams chats
bind the same Legion conversation, the target map keeps the most recent binding;
that is an existing limitation of a one-target-per-conversation map.

### 3.5 Picker card

`src/conversations.ts` exports:

```ts
export const CONVERSATION_PAGE_SIZE = 5;

export function conversationCard(
  entries: ConversationEntry[],
  currentId: string | undefined,
  offset: number,
  key: string,
): IAdaptiveCard;

export function switchedCard(
  entry: ConversationEntry,
  alreadyCurrent?: boolean,
): IAdaptiveCard;

export function conversationErrorCard(message: string): IAdaptiveCard;
```

`conversationCard`:

1. Clamps `offset` to `[0, max(0, entries.length - 1)]`.
2. Takes `entries.slice(offset, offset + CONVERSATION_PAGE_SIZE)`.
3. Builds a card with:
   - `🔀 Switch conversation` (bold heading).
   - `Current: <title>` when the current id is in the list; `Current: <id> (not in
     the active list)` when it is not; `Current: (none)` when no mapping exists.
   - `Showing <start>–<end> of <total>`.
   - A `FactSet` with one fact per page entry: title = the conversation title
     truncated to 40 characters, value = `<n> messages · updated <YYYY-MM-DD>` plus
     ` · current` when applicable.
   - One `Action.Execute` per page entry, titled with the conversation title
     truncated to 40 characters, data:
     ```json
     { "action": "switch", "conversationId": "<id>", "key": "<key>" }
     ```
   - A `More…` action when `offset + 5 < entries.length`, data:
     ```json
     { "action": "switch", "offset": <nextOffset>, "key": "<key>" }
     ```
4. The action list never exceeds six entries: at most five conversation buttons
   plus one `More…`.
5. The empty state is a card with the heading and
   `No conversations yet. Send a message to start one.` and no actions.

`switchedCard` shows `✅ Switched conversation` or
`✅ Already on this conversation`, then a `FactSet` with title, message count, and
updated date, and a line saying future messages in this chat go to that
conversation. It has no actions.

`conversationErrorCard` is a minimal card with `⚠️ <message>` and no actions.

The builder caps titles at 40 characters and pages at five entries. At the default
`maxCardBytes` (28,000) the card is far below the limit; tests assert the bound.
Deployments that lower `maxCardBytes` substantially may need a smaller page size.

### 3.6 `/switch`

**No argument.** Call `listConversations(ctx, participantId)`. If it fails, reply
`Unable to list conversations: <error>` (or the pending/rejected text from
`formatToolResult`). Otherwise send
`conversationCard(entries, state.legionConversations[key], 0, key)`.

**With an argument.** Use the first token as the conversation id.

1. List the participant's active conversations.
2. Find the id. If missing, reply
   `Conversation not found. Use /switch to list active conversations.`
3. If it equals the current id, reply `Already on <title>.`
4. Otherwise call `bindConversation`, then reply
   `Switched to <title> (<n> messages).`

Inbound saves state after `handleCommand` returns `true`, so `/switch` does not
save directly.

With shared fallback (`unknownUserPolicy: 'fallback'`), the chat participant is
the shared fallback participant, so the list may include conversations from other
fallback chats. The card still switches only the current Teams chat's key.

### 3.7 `/compact`

1. If `options.compaction` is absent, reply
   `Compaction is not configured. Set the "compaction.agentId" connector option.`
   and stop.
2. Read `conversationId = state.legionConversations[key]`. If absent, reply
   `Nothing to compact — no conversation yet.`
3. Call `get_conversation` as the participant with `{ conversationId }`.
   - On failure, reply `Compaction failed: <error>`.
   - Take `data.messages` and map to ids. If empty, reply
     `Nothing to compact — conversation is empty.`
4. Choose `instruction = rest.join(' ').trim() || options.compaction.instruction`.
5. Call `compact_conversation` as the participant with:
   ```json
   {
     "conversationId": "<id>",
     "messageIds": ["<oldest>", "..."],
     "agentId": "<options.compaction.agentId>",
     "instruction": "<instruction, omitted when empty>"
   }
   ```
6. On a non-success result, reply with the existing `formatToolResult` text
   (`Error: …`, `Rejected.`, or `Awaiting approval.`), prefixed with
   `Compaction failed: ` for plain errors.
7. On success, read `data.summaryMessageId`. If absent, reply
   `Compacted <n> messages — no summary message was returned.`
8. Call `get_conversation` again and find the message with that id. If its content
   is absent, reply `Compacted <n> messages — summary unavailable.`
9. Stop typing, then reply:
   ```
   🗜 Compacted <n> messages into a summary:

   <summary truncated to 3,500 characters>
   ```

The whole command runs under the command-path typing indicator. `get_conversation`
returns the active chain, which is the contiguous prefix `compact_conversation`
expects.

### 3.8 `/context`

1. Read `conversationId = state.legionConversations[key]`. If absent, reply
   `No conversation yet — send a message first.`
2. Call `get_conversation` as the participant. If it fails, reply
   `Context unavailable: <error>` (or the pending/rejected text from
   `formatToolResult`) and stop.
3. Call `list_conversations` as the participant to find the `ConversationMeta` for
   this id. This supplies `messageCount` and the resolved title. If the list call
   fails or the id is not present, continue with the `get_conversation` title and
   the active-chain length.
4. Call `query_usage` as the participant with `{ conversationId }`.
5. Reply with:

   ```
   📊 Context — <title>
   Conversation: <id>
   Status: <status> · Messages: <total> total, <active> active
   Tokens: in <n> · out <n> · reasoning <n>
   Cache: <n> read / <n> write · Cost: $<n>
   ```

   - `<title>` is the resolved list title, then the `get_conversation` title, then
     the id.
   - `<total>` is `meta.messageCount`; when the meta is unavailable, omit
     `total` and show only `<active> active`.
   - `<active>` is `data.messages.length` from `get_conversation`.
   - Numbers use `toLocaleString('en-US')`; cost uses four decimal places.
   - If `query_usage` fails or returns no totals, replace the last two lines with
     `Usage: unavailable (<error>)` or `Usage: no usage recorded yet.`

The normal output is short and fits in one Teams message; no chunking is needed.

### 3.9 Card action: `card.action.switch`

`bot-pool.ts` registers:

```ts
app.on('card.action.switch', (ctx) => handlers.onConversationAction(entry, ctx as never));
```

`AppHandlers` adds:

```ts
onConversationAction: (entry: BotEntry, ctx: never) => Promise<CardInvokeResponse>;
```

`connector.ts` implements `onConversationAction`:

1. Require the connector to be started (`ctx`, `state`, `store`, `pool`).
2. Check the bot tenant and conversation type with the same rules as
   `onApproval`/`onToolAction`.
3. Resolve the chat participant and key using `resolveChatParticipant`:
   - **Personal:** `findTeamsActor(aadObjectId)`, else the configured fallback
     participant when `unknownUserPolicy === 'fallback'`. Key is
     `${appId}:${participantId}` for a registered user and
     `${appId}:${participantId}:${aadObjectId}` for fallback. This matches
     inbound.
   - **Group:** `findGroupActor(conversationId)`, else the configured fallback
     participant. Key is `${appId}:${participantId}:${conversationId}`.
   - If no participant resolves, return a message response:
     `Not registered.` for personal, `This group is not registered.` for group.
4. Parse `activity.value.action.data`:
   - `action` must be `switch` (the SDK routes by this value; the handler
     re-checks it).
   - `key` — optional string, at most 256 characters. When present it must equal
     the derived key; otherwise use the derived key.
   - `conversationId` — optional string, at most 256 characters, trimmed.
   - `offset` — optional integer. When present, clamp to a non-negative integer.
   - At least one of `conversationId` or `offset` must be present.
   - Invalid payloads return `Invalid switch payload.` without work.
5. Return `cardResponse(runningCard('switch'))` immediately.
6. Schedule the work in the background through `this.queue.run(key, …)`, using the
   same `ACK_SETTLE_MS` delay as tool actions before updating the card.
7. In the background:
   - Call `listConversations(ctx, participantId)`.
   - **More:** if `conversationId` is absent, build
     `conversationCard(entries, currentId, offset, key)` and update the card.
   - **Choose:** if `conversationId` is present, find it in `entries`. If missing,
     show `conversationErrorCard('Conversation not found.')`.
   - If the chosen id is current, show `switchedCard(entry, true)`.
   - Otherwise call `bindConversation`, `await store.save(state)`, and show
     `switchedCard(entry)`.
   - Update the original card via
     `app.api.conversations.updateActivity(teamsConversationId, replyToId, activity)`
     when `replyToId` is present; otherwise send a new card. Use
     `new MessageActivityInput().addCard('adaptive', card)`.
   - On listing or delivery failure, update with
     `conversationErrorCard('Could not switch conversation.')` when possible and
     log a sanitized error.

**Group clickers.** The clicker does not need a Legion identity in a group chat;
the group participant is the actor. The handler only requires the group participant
to resolve. This is intentional: switching rebinds the group's thread and runs no
tool as the human. Personal chats still require the sender to resolve to the chat's
participant (or fallback).

**Queueing.** The switch runs in `SerialQueue` under the derived key. Inbound turns
use the same queue, so an in-flight turn cannot write the old mapping after the
switch. The running card is returned immediately; the queued work waits
`ACK_SETTLE_MS` before updating the card so the ack lands first. If a turn is
already in flight, the switch waits for it and the running card remains until the
turn ends.

### 3.10 Typing

Factor the typing tick in `inbound.ts` into:

```ts
function startTyping(
  context: TeamsContextLike,
  intervalMs: number,
): () => void;
```

- When `intervalMs <= 0`, return a no-op stop function.
- Otherwise send an initial `TypingActivityInput`, start the interval, and return
  an idempotent stop function.

Use it in `runTurn` (replacing the inline tick) and around command handling:

```ts
const stopTyping = startTyping(context, deps.options.typingIntervalMs);
try {
  const handled = await handleCommand({ ..., stopTyping });
  if (handled) { await deps.stateStore.save(deps.state); return; }
} finally {
  stopTyping();
}
```

Apply the same pattern in both the personal and group command paths in
`handleInbound`.

`/compact` calls `deps.stopTyping?.()` before its final reply.

### 3.11 Identity helper

`src/identity.ts` adds:

```ts
export interface ChatParticipant {
  participantId: string;
  key: string;
}

export function resolveChatParticipant(args: {
  collective: CollectiveLike;
  appId: string;
  aadObjectId?: string;
  conversationId?: string;
  isGroup: boolean;
  options: ResolvedOptions;
}): ChatParticipant | undefined;
```

- Personal: `findTeamsActor(aadObjectId)`; else when
  `options.unknownUserPolicy === 'fallback'` use `options.fallbackUserId` if that
  participant is active and not retired.
- Group: `findGroupActor(conversationId)`; else the same fallback rule.
- Build the key exactly as inbound does, including the aadObjectId suffix for
  fallback personal chats.

This helper does not create pending pairs; it only resolves participants that
inbound would already have resolved.

### 3.12 Error handling

- Command failures use plain-text replies and keep state unchanged.
- Card-action failures update the card with a short error card.
- Missing or ungranted command tools surface the core tool error in the reply.
- Card payloads are treated as untrusted: lengths, types, the `key`, and
  conversation membership are all validated.
- No credentials or raw SDK errors are logged; use the existing `describeError`
  sanitization.

## 4. State and compatibility

No state schema change. The connector keeps schema v3. Switching writes the
existing maps:

```ts
state.legionConversations[key] = conversationId;
state.legionConversationTargets[conversationId] = key;
```

`/new` keeps its current cleanup. Existing `/new`, `/pair`, `/whoami`, approval
cards, tool buttons, and the three `teams__*` tools remain unchanged.

Conversation ids are Legion ids, not Teams ids. Switching does not change the Teams
conversation record (`state.conversations[key]`); outbound delivery still routes to
the same Teams chat.

## 5. Options

Add to `TeamsConnectorOptions` and `ResolvedOptions`:

```ts
compaction?: { agentId: string; instruction?: string };
```

Validation in `resolveOptions`:

- `compaction` must be an object when present.
- `agentId` must be a non-empty string.
- `instruction` must be a string when present.

Without `compaction`, `/compact` replies with setup guidance and makes no tool call.

README options table gains:

| Option | Default | Purpose |
| --- | --- | --- |
| `compaction` | unset | `{ agentId, instruction? }` for `/compact`; `agentId` is the summarizer agent. |

## 6. Provisioning and grants

`autoProvision` and `autoProvisionGroup` currently grant `communicate` and
`approval_response`. Extend both tool maps with:

```ts
list_conversations: 'auto',
get_conversation: 'auto',
query_usage: 'auto',
compact_conversation: 'auto',
```

Existing paired users and groups are not auto-provisioned again. README documents a
manual grant snippet for the four tools and notes that commands fail closed when a
grant is missing. The four tools are ordinary Legion tools; no new core grant
mechanism is introduced.

## 7. Manifest

`scripts/provision-teams-bot.mjs` adds:

```json
"commandLists": [
  {
    "scopes": ["personal", "groupChat"],
    "commands": [
      { "title": "new", "description": "Start a new conversation" },
      { "title": "switch", "description": "List and switch to an existing conversation" },
      { "title": "compact", "description": "Compact this conversation" },
      { "title": "context", "description": "Show usage and context details" },
      { "title": "whoami", "description": "Show identity and chat details" },
      { "title": "pair", "description": "Pair a pending Teams identity" },
      { "title": "help", "description": "List commands" }
    ]
  }
]
```

The manifest `version` becomes `1.1.0` so Teams accepts an update. The script still
regenerates `manifest.json` and the zip; the operator uploads or updates the app in
Teams. README and deploy docs describe the reprovisioning step and note that
existing installs do not see the command list until the app is updated.

Teams command titles are lowercase, at most 32 characters; descriptions are at
most 128 characters.

## 8. Testing

- `commands.test.ts`
  - `/switch` no argument sends a card; direct switch binds and confirms; current
    id is a no-op; unknown id reports not found; listing failure reports the error.
  - Group and fallback keys bind the correct entry and clean the old target.
  - `/compact` without `compaction` makes no tool call; empty conversation is a
    no-op; success reports the summary; instruction overrides the default; tool
    errors are reported.
  - `/context` shows title, counts, and totals; a missing meta degrades to active
    count; usage failure degrades gracefully; no conversation reports guidance.
  - `/help` includes the new commands and group scoping.
- `conversations.test.ts` (new)
  - Card shape, five-per-page, `More…` offset, current marker, empty state,
    truncation, title fallback, and size bound.
  - `bindConversation` cleans the previous target only when it still points at the
    key; rebinding the same id is a no-op; rebinding to an id already targeted by
    another key updates the target.
  - `listConversations` maps `ConversationMeta`, sorts newest first, and throws on
    tool failure.
- `connector.test.ts`
  - Switch choice updates state and the card; `More…` pages without changing
    state; an unowned/invalid conversation shows the error card; group and fallback
    keys resolve; the queue serializes a switch after an in-flight turn; a missing
    `replyToId` sends a follow-up card.
  - Unknown personal clicker, unknown group, tenant mismatch, and unsupported
    conversation type return message responses.
- `bot-pool.test.ts` — `card.action.switch` registration and handler dispatch.
- `options.test.ts` — `compaction` validation.
- `provision.test.ts` — manifest `commandLists` and version `1.1.0`.
- `identity.test.ts` — `resolveChatParticipant` personal/group/fallback keys.
- Live check: in a personal chat, create two conversations, run `/switch`, choose
  the older one, send a message, and verify it lands in the chosen conversation;
  repeat in a group chat. Run `/context` and `/compact` against a granted
  participant.

## 9. Follow-ups

- Team channel support, including @mention routing and channel command lists.
- A `Show archived` toggle on the picker.
- Per-chat conversation aliases and pinned conversations.
- A `New conversation` button on the picker.
- Server-side card state for tamper-proof switching and idempotency.
- Branch switching (`switch_branch`) from Teams.
