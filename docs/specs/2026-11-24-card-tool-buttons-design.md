# Legion Teams Connector — Tool Buttons on Adaptive Cards

**Date:** 2026-11-24
**Status:** Approved for planning
**Package:** `@legion-collective/connector-teams`
**Builds on:** `docs/specs/2026-11-21-teams-connector-design.md`

## 1. Context

The Teams connector can send Adaptive Cards (`teams__send_card`) and already handles
Adaptive Card actions for approvals: approval cards use `Action.Execute` buttons whose
`data.action` is `approve` or `reject`, and the connector routes those invokes through
`card.action.approve` / `card.action.reject` to `approval_response`.

This spec generalizes that mechanism. An agent can attach **tool buttons** to a card it
sends. Clicking a button runs a Legion tool as the **clicking user**, then updates the
card with the result.

The clicker is the actor. This keeps the existing fail-closed model: `ConnectorContext.callTool`
authorizes the participant's own grants, argument scopes, and tool visibility before
execution. An agent cannot use a button to grant a user a tool the user does not have.

### Goals

- `teams__send_card` accepts an optional `buttons` array.
- A button names a Legion tool and optional fixed args.
- Clicking resolves the Teams user to a Legion participant and runs the tool as them.
- The card is acknowledged immediately, then updated in place with the result.
- Results can use an agent-supplied Adaptive Card template or a generic renderer.
- Unknown or ungranted users fail closed.
- Existing cards without buttons behave exactly as before.
- No Legion core changes.

### Non-goals

- Dynamic form fields (`Input.*` values merged into args at click time).
- Approval-gated button tools. V1 reports that approval is required; the approval flow
  is a follow-up.
- Server-side button storage or idempotency keys.
- Tool-provided result cards (`result.data.card`).
- Buttons in team channels.

## 2. Decisions

| Decision | Choice | Rationale |
| --- | --- | --- |
| Authoring surface | Optional `buttons: [{title, tool, args?, resultCard?}]` on `teams__send_card` | Simple for agents; connector builds well-formed actions |
| Button action | `Action.Execute` with `data.action = 'tool'` | Matches the SDK's `card.action.<data.action>` routing and the existing approval pattern |
| Button limit | 6 total actions per card | Adaptive Cards renders at most six actions |
| Actor | The clicking Teams user | `ctx.callTool` enforces the clicker's grants and scopes; fail-closed |
| Unknown clicker | Reject with “Not registered.” | Mirrors `onApproval`; no auto-provisioning on click |
| Ack | Card response replacing the card with “Running <tool>…” and no actions | Fast invoke response; prevents double-clicks |
| Completion | Update the same card in place | Keeps the result next to the card that produced it |
| Update target | `activity.replyToId` | Teams sets it to the card message id for `adaptiveCard/action` invokes |
| Fallback | Send the result as a follow-up message | Covers invokes without `replyToId` |
| Result formatting | Agent `resultCard` template, else generic renderer | Gives control when it matters, readable output otherwise |
| Templating | `adaptivecards-templating` + `adaptive-expressions` | Official Adaptive Card data binding; supports `${...}` and `$data` |
| Template context | `{status, data, error, message, tool, args, clicker}` | Everything a formatter needs; `clicker` is `{id, name?}` |
| Conversation | Reuse the mapped Legion conversation; otherwise an ephemeral call | Avoids creating a Legion conversation per click |
| State | No schema change; button data travels in the card action payload | V1 stays small; server-side storage is a follow-up |
| Options | No new option; buttons ride under `tools: ['cards']` | `teams__send_card` already controls card access |

## 3. Architecture

### Send path

`teams__send_card` validates `buttons`, builds one `Action.Execute` per button, appends
them to `card.actions`, and enforces the combined action limit and `maxCardBytes` on the
outgoing card. When `buttons` is absent, the card is sent unchanged.

Button action payload:

```json
{
  "action": "tool",
  "tool": "teams__post_reaction",
  "args": { "messageId": "m-1", "reaction": "like" },
  "resultCard": { "type": "AdaptiveCard", "version": "1.4", "body": [] }
}
```

`args` and `resultCard` are omitted when the agent does not supply them.

### Click path

Each bot registers `card.action.tool`. The connector's handler:

1. Checks the bot tenant and conversation type using the same rules as approvals.
2. Resolves the clicker with `readTeamsId(activity.from)` and `findTeamsActor`.
3. Validates `tool` and `args` from the action payload.
4. Resolves the Legion conversation for this Teams conversation, if one is known.
5. Returns a card response with a running card.
6. Runs `ctx.callTool(clicker.id, tool, args ?? {}, {conversationId})` in the background.
7. Builds the result card and updates the card message (`replyToId`) in place, or sends a
   follow-up message when `replyToId` is absent.

The clicker's participant id is the only authority used. `callTool` checks the tool
registry, the participant's `tools` map, and argument scopes.

### Result formatting

`resultCard` is an Adaptive Card template. The connector expands it with
`adaptivecards-templating` against:

```ts
{
  status: ToolResult['status'],
  data?: unknown,
  error?: string,
  message?: string,
  tool: string,
  args?: unknown,
  clicker?: { id: string; name?: string },
}
```

The root is `$root`, so `${status}`, `${data.reaction}`, `${clicker.name}`, and
`$data: "${data.conversations}"` work. If expansion fails or produces a non-card, the
connector uses the generic renderer.

The generic renderer:

- Header: success, error, pending approval, or rejected.
- Error/message text when present.
- Strings as a TextBlock.
- Flat scalar object fields as a FactSet.
- Array properties rendered item by item (objects as FactSets, primitives as bullets).
- Unknown or nested-only data as a truncated monospace JSON block.
- A compact fallback when the card exceeds `maxCardBytes`.

### Failure behavior

- Missing tool, malformed args, unknown clicker, tenant mismatch, or unsupported
  conversation type return a message response and do not call the tool.
- A `requires_approval` result renders a card saying the tool needs approval. V1 does not
  create an approval request from a button click.
- Tool exceptions render an error card.
- Card update failures are logged; delivery already happened.

## 4. State and compatibility

No state schema change. The button's tool, args, and template travel in the Adaptive Card
action payload. The existing `legionConversations` map is used to reuse a Legion
conversation when one exists; otherwise `callTool` receives an empty conversation id and
uses an ephemeral thread.

Existing `teams__send_card` behavior, approval cards, state schema v3, and the three
`teams__*` tools remain unchanged.

## 5. Testing

- Unit tests for the generic renderer, template expansion, size fallback, button action
  construction, `buttons` validation, event registration, and the connector handler.
- Connector handler tests cover tenant mismatch, conversation type, unknown clicker,
  missing/invalid payload, success, `requires_approval`, tool failure, mapped
  conversation reuse, and the follow-up path when `replyToId` is absent.
- Live test: send a card with a `teams__post_reaction` button, click it as a registered
  user, and verify the reaction and the updated card.

## 6. Follow-ups

- Approval-gated button tools via a core `ConnectorContext` approval seam.
- Server-side button records for tamper-proof args and idempotency.
- Dynamic form fields and `Input.*` value binding.
- Tool-provided result cards (`result.data.card`).
