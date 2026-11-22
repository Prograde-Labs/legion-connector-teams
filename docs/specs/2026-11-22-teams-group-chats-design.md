# Legion Teams Connector — Group Chats as Participants

**Date:** 2026-11-22
**Status:** Approved for planning
**Package:** `@legion-collective/connector-teams`
**Builds on:** `docs/specs/2026-11-21-teams-connector-design.md`

## 1. Context

The first version of the Teams connector supports personal 1:1 chats only. Its
non-goals listed group chats as a follow-up, with the intended shape: "the group
`conversationId` becomes a Legion user participant."

This spec adds that follow-up for **group chats** (`conversationType === 'groupChat'`).
Team channels remain a separate follow-up because they need @mention routing, RSC
permissions, channel installation, and threaded replies.

A group chat is modelled as **one Legion participant**, not as a conversation between
several human participants. Legion's `communicate` tool is a 1:1 exchange between a
sender and a recipient; making the group the sender is what lets the agent reply into
the group without inventing a multi-party runtime. The humans in the group are
represented by attribution text, not separate Legion participants.

### Goals

- A Teams group chat can @mention a Legion bot and receive a reply in that group.
- The group is a Legion `user` participant with its own memory and conversation state.
- Unknown groups use the same provisioning policy as unknown personal users.
- The agent sees who spoke via a configurable sender prefix.
- Native streaming, typing, tool cards, and approvals work in group chats.
- Any registered group member with `approval_response` authority can decide an
  approval card, and the decided card names the decider.
- Existing personal-chat behavior and state remain compatible.

### Non-goals

- Team channels (`conversationType === 'channel'`), channel threads, and RSC.
- Per-human Legion participants inside a group chat.
- Member rosters, targeted messages, meeting chats, and channel @mention handling.
- File attachments.
- Multiple bots answering in the same group chat; each bot still answers only when
  it is the one @mentioned.

## 2. Decisions

| Decision | Choice | Rationale |
| --- | --- | --- |
| Representation | One `user` participant per group chat | Matches Legion's sender/recipient `communicate` model; no core runtime change |
| Identity | `{connector:'teams-group', externalId:'<conversation.id>'}` | Keeps personal (`teams`) and group identities in separate namespaces; `conversation.id` is stable for group chats |
| Provisioning | Same `unknownUserPolicy` as personal chats | One trust model; `auto-provision` works in the local deployment |
| Participant ID | `teams-group-<slug(name)>-<hash6(conversationId)>` | Stable, human-readable, collision-resistant |
| Attribution | Configurable template, default `[{name}]: {message}` | The agent needs speaker context; deployments may prefer another format |
| Streaming | Native streaming, with final-send fallback | Teams released streaming for group chats on 2026-10-28; `@microsoft/teams.apps` 2.2.1 sets `replyToId` |
| Concurrency | Serialize turns per conversation | Teams allows one open stream per chat |
| Approvals | Any registered member with authority | AuthEngine already checks the clicking participant's `approvalAuthority` against the requester |
| Commands | Group-scoped, except `/pair` as the typing human | `/new` and `/whoami` describe the group; `/pair` must use the human's authority |
| State | Connector-owned schema v2 with migration | No core changes; old state files keep working |
| Outbound routing | Prefer a stored record; add a Legion-conversation target map | Replies must land in the exact group, including fallback groups |

## 3. Architecture

```
Teams group chat
   │  activity.conversation.conversationType = 'groupChat'
   ▼
TeamsConnector
   ├── Inbound
   │     ├── tenant check
   │     ├── strip only the bot's @mention
   │     ├── resolve group by conversation.id (teams-group identity)
   │     ├── unknown policy → pairing / auto-provision / reject / fallback
   │     ├── sender attribution prefix
   │     └── communicate from the group participant to the mentioned bot
   ├── Approvals   clicking human resolved via their teams identity
   ├── Outbound    target map → exact group conversation
   └── State       conversations, pendingPairs, legionConversations, targets
   ▼
Legion core: ConnectorContext.streamTool / callTool / registry (unchanged)
```

Personal chats keep the existing path. The connector decides the path from
`activity.conversation.conversationType`; no core code changes.

## 4. Identity

### 4.1 Group identity

A group chat participant carries:

```json
{ "connector": "teams-group", "externalId": "<conversation.id>" }
```

`conversation.id` is the group chat's stable conversation ID (for example
`19:...@thread.v2`). It is the same value used by `app.send` and by the activity's
`conversation.id`.

Personal participants keep `{ "connector": "teams", "externalId": "<aadObjectId>" }`.

### 4.2 Resolution

For a `groupChat` activity:

1. `collective.findByIdentity('teams-group', conversation.id)`.
2. A retired participant is treated as unknown.
3. Otherwise apply `unknownUserPolicy` (section 6).

The sender (`activity.from.aadObjectId`) is **not** used to resolve the group. It is
used for attribution and for `/pair` authority.

### 4.3 Attribution

The agent receives group messages with the sender's name prefixed. The template is
configured by `groupSenderPrefix`:

- default: `[{name}]: {message}`
- `null` disables attribution
- `{message}` is required; `{name}` is optional
- `resolveOptions` throws if the template contains no `{message}`

Example: `[Chris Willingham]: can you review this?`

Commands are handled before the prefix is applied.

## 5. Inbound flow

Per-App handlers are unchanged. The message handler branches on conversation type.

### 5.1 Order

1. Reject non-`personal`/non-`groupChat` conversations with
   `Only personal chats and group chats are supported.` (unless `groupChats` is
   disabled, in which case groups get `Only personal chats are supported.`).
2. Enforce the bot tenant (unchanged).
3. Read the sender's Teams ID (unchanged).
4. Strip only the bot's @mention:
   `stripMentionsText(activity, { accountId: activity.recipient.id })`.
   Human-to-human @mentions are preserved.
5. Resolve or provision the group participant.
6. Handle commands.
7. Apply the attribution prefix.
8. Run `communicate` as the group participant.

A bare @mention (empty text after stripping) gets `How can I help?`.

### 5.2 Conversation record

For groups, upsert:

```
<appId>:<participantId>:<conversationId>
  conversationId, conversationType: 'groupChat', name,
  tenantId, serviceUrl, externalId: conversationId, updatedAt
```

Personal records keep `<appId>:<participantId>` with
`conversationType: 'personal'` and `externalId: <aadObjectId>`.

`ctx.registry.setActive(participantId, 'teams')` is called on every inbound
message, as before.

### 5.3 Group installs

`install.add` and `conversationUpdate` for group chats are ignored for state. The
first inbound @mention creates the conversation record. Proactive messages before
that first mention are best-effort (section 10).

## 6. Provisioning

`groupChats` (default `true`) enables the group path. The same `unknownUserPolicy`
applies:

| Policy | Group behavior |
| --- | --- |
| `pairing` (default) | Reply with a 6-character code bound to `(appId, 'teams-group', conversationId)`; a registered human runs `/pair <code> <participant-id>` |
| `auto-provision` | `create_user` then `pair` with connector `teams-group` |
| `reject` | Reply with the conversation ID and a registration hint |
| `fallback` | Use `fallbackUserId`; conversation state is isolated per group conversation ID |

### 6.1 Auto-provision

1. `id = teams-group-<slug(conversation.name ?? 'group')>-<hash6(conversationId)>`
2. `name = conversation.name ?? 'Teams group <hash6>'`
3. `ctx.callTool(provisionerId, 'create_user', { id, name, tools })` where `tools`
   grants `communicate: 'auto'` and `approval_response: 'auto'` (same as personal)
4. `ctx.callTool(pairingParticipantId ?? provisionerId, 'pair', { participantId: id, connector: 'teams-group', externalId: conversationId })`

A `pending_approval` result is reported to the group; an error leaves the group
unresolved. The `create_user`/`pair` flow is unchanged from personal provisioning,
only the connector and external ID differ.

### 6.2 Pairing codes

`pendingPairs` becomes:

```json
{ "<code>": { "appId": "…", "connector": "teams-group", "externalId": "…", "expiresAt": "…" } }
```

Codes remain 6 characters, single-use, TTL-bound, and reused for the same
`(appId, connector, externalId)`.

The unknown-group reply:

> This group isn't registered with Legion yet. Ask someone with the `pair` tool to run
> `/pair <code> <participant-id>`. Group ID: `<conversationId>`.

### 6.3 Retired participants

A retired group participant is treated as unknown and follows the policy.

## 7. Commands

| Command | Group behavior |
| --- | --- |
| `/new` | Clears the group's Legion conversation (`<appId>:<participantId>:<conversationId>`) |
| `/whoami` | Shows the group participant, group conversation ID, sender, and bot app |
| `/pair <code> <participant-id>` | Runs as the typing human when that human is a registered participant; uses the pending `connector`/`externalId` |
| `/help` | Lists commands, including the group behavior |

An unregistered human in a group gets `Only registered users can redeem pairing
codes.` A group participant is not required to have `pair`; the human's authority is
what AuthEngine checks.

## 8. Streaming and typing

### 8.1 Native streaming

Teams supports streaming in group chats as of 2026-10-28. With
`@microsoft/teams.apps` 2.2.1, `HttpStream` sets `replyToId` from the inbound
activity, which is required outside 1:1 chats.

The connector keeps its current chunk mapping:

- `stream.update('Thinking…')` for informative updates
- `stream.emit(delta)` for text
- `stream.close()` for the final message

If `context.stream` is absent, or the stream throws, the connector sends the final
text with a normal message activity.

### 8.2 One stream per chat

Teams allows one open stream per conversation. Inbound turns are serialized per
conversation key with a small promise chain, so two simultaneous @mentions do not
open competing streams. The chain is keyed by `<appId>:<participantId>:<conversationId>`
for groups and `<appId>:<participantId>` for personal chats.

### 8.3 Typing

The periodic typing indicator and "stop typing before the final reply" behavior are
unchanged. Typing activities do not need `replyToId`.

## 9. Approvals

Approval cards are posted in the group chat by the existing `communicate` path.

When a member clicks Approve/Reject:

1. Resolve the clicking human with `collective.findByIdentity('teams', aadObjectId)`.
   An unregistered clicker gets `Not registered.`
2. Call `approval_response` as that human. AuthEngine checks their
   `approvalAuthority` against the requester (normally the assistant). No core
   changes are needed.
3. Return a decided Adaptive Card with no buttons: outcome, tool, args, requester,
   optional rejection reason, and the decider's display name.
4. Deliver any resumed replies and successor approval cards back to the group.
5. On failure, keep the buttons and return an error message.

The card input `reason` continues to be merged into `Action.Execute` data and passed
as `message` to `approval_response`.

## 10. Outbound

`deliver(message)`:

1. Resolve the bot from the sender's `teams-bot` identity (unchanged).
2. If `legionConversationTargets[message.conversationId]` exists, use that
   conversation record.
3. Else, if the recipient has a `teams-group` identity, send directly to
   `externalId` (the group conversation ID). If no record exists, store one under
   `<appId>:<participantId>:<externalId>` with `conversationType: 'groupChat'`,
   `externalId`, and `updatedAt`.
4. Else use the existing personal path: stored record, then create a personal
   conversation from the recipient's `teams` identity.
5. Chunk the content, add the AI label, and send. 403 handling is unchanged.

The target map removes the "most recently updated record" ambiguity for fallback
groups and for a participant paired to several group identities.

## 11. Persistent state

State schema v2:

```json
{
  "version": 2,
  "conversations": {
    "<appId>:<participantId>": {
      "conversationId": "…",
      "conversationType": "personal",
      "externalId": "…",
      "tenantId": "…",
      "serviceUrl": "…",
      "displayName": "…",
      "updatedAt": "…"
    },
    "<appId>:<participantId>:<conversationId>": {
      "conversationId": "…",
      "conversationType": "groupChat",
      "externalId": "…",
      "name": "…",
      "tenantId": "…",
      "serviceUrl": "…",
      "updatedAt": "…"
    }
  },
  "pendingPairs": {
    "<code>": { "appId": "…", "connector": "teams", "externalId": "…", "expiresAt": "…" }
  },
  "legionConversations": {
    "<appId>:<participantId>[:<conversationId>]": "<legion-conversation-id>"
  },
  "legionConversationTargets": {
    "<legion-conversation-id>": "<platform-conversation-key>"
  }
}
```

### 11.1 Migration

On load:

- v1 `pendingPairs.aadObjectId` becomes `externalId` with `connector: 'teams'`.
- v1 conversation records gain `conversationType: 'personal'` and
  `externalId` from `aadObjectId` (when present).
- `legionConversationTargets` starts empty.
- The file is rewritten as v2 on the next mutation.

Corrupt files are still backed up to `state.json.bak-<timestamp>` and replaced with
empty v2 state. Writes remain atomic and mode 0600; no secrets are stored.

## 12. Configuration

| Option | Type | Default | Meaning |
| --- | --- | --- | --- |
| `groupChats` | `boolean` | `true` | Enable the group-chat path |
| `groupSenderPrefix` | `string \| null` | `'[{name}]: {message}'` | Attribution template; `null` disables |

All existing options are unchanged. `resolveOptions` validates `groupSenderPrefix`
and throws if it lacks `{message}`.

## 13. Testing

Unit tests (no live Teams required):

- Group resolution for `pairing`, `auto-provision`, `reject`, and `fallback`.
- Retired group participant treated as unknown.
- Unnamed group naming and ID generation.
- Mention stripping: only the bot's mention is removed; human @mentions survive.
- Attribution template: default, custom, disabled, and missing `{message}`.
- Commands: `/new` and `/whoami` on the group key; `/pair` as registered and
  unregistered humans; pending connector/externalId used.
- Approvals in a group: clicker resolution, decided card with reason and decider,
  resumed delivery, failure keeps buttons.
- Outbound: target-map routing, direct group send, stored record preference,
  fallback group isolation.
- Streaming: native stream used when present; final-send fallback when absent or on
  error; per-conversation serialization.
- State v1 → v2 migration and corrupt-file handling.
- Existing personal-chat tests remain green.

Verification: full test suite, typecheck, build, and a live retest in a real Teams
group chat.

## 14. Risks and open items

| Risk | Mitigation |
| --- | --- |
| Bots receive only @mentions in group chats by default | Respond to every received group message; mention stripping is idempotent |
| One open stream per chat | Serialize turns per conversation key |
| Concurrent @mentions in one group | Per-conversation queue; later turns wait |
| Group install before first mention | State is created on first inbound message; proactive sends are best-effort |
| A participant paired to several groups | Group conversation keys include the conversation ID; target map routes replies |
| Fallback groups sharing one participant | Target map keeps replies in the originating group |
| Unnamed group chats | Use `Teams group <hash6>` |
| Channels need more platform work | Explicit non-goal; same participant model can be reused later |

## 15. Future work

- Team channels as participants, with @mention routing, RSC permissions, and threads.
- Member roster awareness and per-human attribution metadata.
- Targeted messages (private replies inside a group).
- Channel install and Graph-based membership discovery.
