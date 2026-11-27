# Teams Conversation Commands and Picker Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add `/switch`, `/compact`, and `/context` to the Teams connector, including an Adaptive Card conversation picker, compaction configuration, auto-provisioned tool grants, and a Teams app manifest command list.

**Architecture:** `handleCommand` gains three cases backed by a new `src/conversations.ts` module for listing, binding, and card construction. `/switch` sends an Adaptive Card with `Action.Execute` buttons; a new `card.action.switch` handler acks with a running card, lists the chat participant's active conversations, rebinds the chat under the per-chat `SerialQueue`, and updates the card in place. `/compact` calls `compact_conversation` through a configured summarizer agent; `/context` combines `get_conversation`, `list_conversations`, and `query_usage`. No Legion core changes and no state schema change.

**Tech Stack:** TypeScript ESM/NodeNext, Node 22, `@microsoft/teams.apps` 2.2.1, `@microsoft/teams.api` 2.2.1, `@microsoft/teams.cards` 2.2.1, Vitest.

**Spec:** `docs/specs/2026-11-25-teams-commands-design.md`

## Global Constraints

- **No Legion core changes.** Use `ConnectorContext.callTool(participantId, toolName, args, opts?)` with the existing core tools `list_conversations`, `get_conversation`, `query_usage`, and `compact_conversation`.
- **State schema stays v3.** Switching rebinds `state.legionConversations[key]` and updates `state.legionConversationTargets[conversationId]`.
- **Commands act as the chat participant.** Personal chat → the paired user (including the fallback suffix); group chat → the group participant.
- **Card payloads are untrusted.** Validate `action`, `key`, `conversationId` (≤ 256 chars), and `offset` (finite, clamped). A chosen id must appear in the participant's active list.
- **At most six Adaptive Card actions per card:** five conversation buttons plus one `More…`.
- **Running-card ack first.** Return `cardResponse(runningCard('switch'))` immediately, then update in place after `ACK_SETTLE_MS`.
- **Queue key must match inbound.** `onMessage` queues by Teams chat (`${appId}:${aadObjectId}` personal, `${appId}:${teamsConversationId}` group). The switch handler must use the same key so an in-flight turn cannot write the old mapping after a switch.
- **Unknown slash commands remain ordinary messages** (`handleCommand` returns `false`).
- **No automatic grants for existing participants.** Auto-provisioning adds the four tools; README documents the manual grant.
- **Manifest version `1.1.0`** and a `commandLists` entry scoped to `personal` and `groupChat`.
- Every task ends with its tests passing, `npm run typecheck` clean, and a commit.
- Connector test command: `npm test` (Vitest). Single file: `npx vitest run src/<file>.test.ts`. Build: `npm run build`.
- Work in `threads/2026-11-25-teams-commands/legion-connector-teams` on branch `feat/teams-commands`.

## File Structure

**Created**
- `src/conversations.ts` — conversation list mapping, binding, picker/result cards, error text.
- `src/conversations.test.ts` — list/bind/card tests.

**Modified**
- `src/options.ts`, `src/options.test.ts` — `compaction` option and validation.
- `src/identity.ts`, `src/identity.test.ts` — `resolveChatParticipant`.
- `src/commands.ts`, `src/commands.test.ts` — `/switch`, `/compact`, `/context`, help text, `sendCard`/`stopTyping`.
- `src/inbound.ts`, `src/inbound.test.ts` — `startTyping`, `sendCard` wiring, auto-provision grants.
- `src/bot-pool.ts`, `src/bot-pool.test.ts` — `onConversationAction`, `card.action.switch`.
- `src/connector.ts`, `src/connector.test.ts` — switch handler, queue-key helper.
- `scripts/provision-teams-bot.mjs`, `src/provision.test.ts` — manifest command list and version.
- `README.md`, `docs/deploy.md` — commands, option, grants, reprovisioning.

## Setup

- [ ] Verify the worktree, branch, and spec commit:

```bash
cd threads/2026-11-25-teams-commands/legion-connector-teams
git status --short
git branch --show-current
git log --oneline -1
```

Expected: clean tree, `feat/teams-commands`, and `c6551a6 docs(teams): specify conversation commands and picker`.

---

### Task 1: `compaction` connector option

**Files:**
- Modify: `src/options.ts`
- Test: `src/options.test.ts`

**Interfaces:**
- Produces: `CompactionOptions { agentId: string; instruction?: string }`
- Produces: `TeamsConnectorOptions.compaction?: CompactionOptions` and `ResolvedOptions.compaction?: CompactionOptions`

- [ ] **Step 1: Write the failing tests**

Append to `src/options.test.ts`:

```ts
describe('resolveOptions compaction', () => {
  it('defaults to unset', () => {
    expect(resolveOptions({ stateDir: '/tmp/x' }).compaction).toBeUndefined();
  });

  it('accepts an agent and instruction', () => {
    expect(resolveOptions({
      stateDir: '/tmp/x',
      compaction: { agentId: 'summarizer', instruction: 'Be brief.' },
    }).compaction).toEqual({ agentId: 'summarizer', instruction: 'Be brief.' });
  });

  it('accepts an agent without an instruction', () => {
    expect(resolveOptions({
      stateDir: '/tmp/x',
      compaction: { agentId: 'summarizer' },
    }).compaction).toEqual({ agentId: 'summarizer' });
  });

  it.each([
    [{ agentId: '' }, /agentId/],
    [{ agentId: '   ' }, /agentId/],
    [{ agentId: 'summarizer', instruction: 42 }, /instruction/],
  ])('rejects invalid compaction %j', (compaction, message) => {
    expect(() => resolveOptions({ stateDir: '/tmp/x', compaction: compaction as never }))
      .toThrow(message);
  });

  it('rejects a non-object compaction value', () => {
    expect(() => resolveOptions({ stateDir: '/tmp/x', compaction: 'summarizer' as never }))
      .toThrow(/compaction/);
  });
});
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `npx vitest run src/options.test.ts`
Expected: FAIL — `compaction` is not in `ResolvedOptions`, and invalid values are not rejected.

- [ ] **Step 3: Implement the option**

In `src/options.ts`, add the interface and fields:

```ts
export interface CompactionOptions {
  agentId: string;
  instruction?: string;
}
```

Add `compaction?: CompactionOptions;` to `TeamsConnectorOptions` and `compaction?: CompactionOptions;` to `ResolvedOptions`.

Add the resolver before `resolveOptions`:

```ts
function resolveCompaction(
  value: TeamsConnectorOptions['compaction'],
): CompactionOptions | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    throw new Error('compaction must be an object with an agentId');
  }
  if (typeof value.agentId !== 'string' || !value.agentId.trim()) {
    throw new Error('compaction.agentId must be a non-empty string');
  }
  if (value.instruction !== undefined && typeof value.instruction !== 'string') {
    throw new Error('compaction.instruction must be a string');
  }
  return {
    agentId: value.agentId.trim(),
    ...(value.instruction === undefined ? {} : { instruction: value.instruction }),
  };
}
```

Add `compaction: resolveCompaction(options.compaction),` to the object returned by `resolveOptions`.

- [ ] **Step 4: Run the tests and typecheck**

Run: `npx vitest run src/options.test.ts && npm run typecheck`
Expected: PASS, no type errors.

- [ ] **Step 5: Commit**

```bash
git add src/options.ts src/options.test.ts
git commit -m "feat(teams): add compaction connector option"
```

---

### Task 2: Conversation listing and binding helpers

**Files:**
- Create: `src/conversations.ts`
- Create: `src/conversations.test.ts`

**Interfaces:**
- Produces: `ConversationEntry { id: string; title: string; messageCount: number; updatedAt: string; status: 'active' | 'archived' }`
- Produces: `ConversationToolError extends Error` with `result: { status: string; error?: string; message?: string }`
- Produces: `listConversations(ctx: ConnectorContext, participantId: string): Promise<ConversationEntry[]>`
- Produces: `bindConversation(state: TeamsState, key: string, conversationId: string): void`
- Produces: `conversationErrorText(error: unknown): string`

- [ ] **Step 1: Write the failing tests**

Create `src/conversations.test.ts`:

```ts
// src/conversations.test.ts
import { describe, expect, it, vi } from 'vitest';
import type { ConnectorContext } from '@legion-collective/core';
import {
  ConversationToolError,
  bindConversation,
  conversationErrorText,
  listConversations,
} from './conversations.js';
import { emptyState } from './state.js';

function ctx(result: unknown): ConnectorContext {
  return {
    callTool: vi.fn(async () => ({ result, conversationId: 'c' })),
  } as unknown as ConnectorContext;
}

describe('listConversations', () => {
  it('calls list_conversations as the participant and maps active conversations', async () => {
    const context = ctx({
      status: 'success',
      data: {
        conversations: [
          {
            id: 'older',
            resolvedTitle: 'Older',
            messageCount: 2,
            updatedAt: '2026-11-24T10:00:00.000Z',
            status: 'active',
          },
          {
            id: 'newer',
            sharedTitle: 'Newer shared',
            messageCount: 5,
            updatedAt: '2026-11-25T10:00:00.000Z',
            status: 'active',
          },
        ],
      },
    });
    const entries = await listConversations(context, 'alice');
    expect(context.callTool).toHaveBeenCalledWith('alice', 'list_conversations', {
      participantId: 'alice',
      status: 'active',
    });
    expect(entries.map((entry) => entry.id)).toEqual(['newer', 'older']);
    expect(entries[0]).toMatchObject({
      title: 'Newer shared',
      messageCount: 5,
      updatedAt: '2026-11-25T10:00:00.000Z',
      status: 'active',
    });
  });

  it('falls back through resolved, shared, title, then id', async () => {
    const context = ctx({
      status: 'success',
      data: {
        conversations: [
          { id: 'a', title: 'Plain title', messageCount: 1, updatedAt: '2026-11-25T00:00:00.000Z' },
          { id: 'b', messageCount: 1, updatedAt: '2026-11-25T00:00:00.000Z' },
        ],
      },
    });
    const entries = await listConversations(context, 'alice');
    expect(entries.map((entry) => entry.title)).toEqual(['Plain title', 'b']);
  });

  it('treats a missing conversation list as empty', async () => {
    await expect(listConversations(ctx({ status: 'success', data: {} }), 'alice'))
      .resolves.toEqual([]);
  });

  it('throws a typed error on tool failure', async () => {
    const context = ctx({ status: 'error', error: 'not authorized' });
    const error = await listConversations(context, 'alice').catch((value) => value);
    expect(error).toBeInstanceOf(ConversationToolError);
    expect((error as ConversationToolError).result).toEqual({
      status: 'error',
      error: 'not authorized',
    });
  });
});

describe('conversationErrorText', () => {
  it('uses the plain error for error results', () => {
    expect(conversationErrorText(new ConversationToolError({
      status: 'error',
      error: 'not authorized',
    }))).toBe('not authorized');
  });

  it('uses the approval text for pending results', () => {
    expect(conversationErrorText(new ConversationToolError({ status: 'pending_approval' })))
      .toBe('Awaiting approval.');
  });

  it('uses the rejection message', () => {
    expect(conversationErrorText(new ConversationToolError({
      status: 'rejected',
      message: 'Denied by policy.',
    }))).toBe('Denied by policy.');
  });

  it('falls back to the thrown message', () => {
    expect(conversationErrorText(new Error('network down'))).toBe('network down');
  });
});

describe('bindConversation', () => {
  it('binds a new conversation and cleans the previous target', () => {
    const state = emptyState();
    state.legionConversations['app:alice'] = 'old';
    state.legionConversationTargets['old'] = 'app:alice';

    bindConversation(state, 'app:alice', 'new');

    expect(state.legionConversations['app:alice']).toBe('new');
    expect(state.legionConversationTargets['new']).toBe('app:alice');
    expect(state.legionConversationTargets['old']).toBeUndefined();
  });

  it('does nothing when the conversation is already current', () => {
    const state = emptyState();
    state.legionConversations['app:alice'] = 'same';
    state.legionConversationTargets['same'] = 'app:alice';

    bindConversation(state, 'app:alice', 'same');

    expect(state.legionConversations['app:alice']).toBe('same');
    expect(state.legionConversationTargets['same']).toBe('app:alice');
  });

  it('keeps the previous target when another chat now owns it', () => {
    const state = emptyState();
    state.legionConversations['app:alice'] = 'shared';
    state.legionConversationTargets['shared'] = 'app:bob';

    bindConversation(state, 'app:alice', 'new');

    expect(state.legionConversationTargets['shared']).toBe('app:bob');
  });
});
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `npx vitest run src/conversations.test.ts`
Expected: FAIL — `./conversations.js` does not exist.

- [ ] **Step 3: Implement `src/conversations.ts`**

```ts
// src/conversations.ts
import type { ConnectorContext } from '@legion-collective/core';
import { formatToolResult } from './render.js';
import type { TeamsState } from './state.js';

export interface ConversationEntry {
  id: string;
  title: string;
  messageCount: number;
  updatedAt: string;
  status: 'active' | 'archived';
}

interface ConversationMetaLike {
  id?: unknown;
  resolvedTitle?: unknown;
  sharedTitle?: unknown;
  title?: unknown;
  messageCount?: unknown;
  updatedAt?: unknown;
  status?: unknown;
}

export class ConversationToolError extends Error {
  readonly result: { status: string; error?: string; message?: string };

  constructor(result: { status: string; error?: string; message?: string }) {
    super(result.error ?? result.message ?? result.status);
    this.name = 'ConversationToolError';
    this.result = result;
  }
}

function text(value: unknown): string | undefined {
  return typeof value === 'string' && value.trim() ? value.trim() : undefined;
}

export async function listConversations(
  ctx: ConnectorContext,
  participantId: string,
): Promise<ConversationEntry[]> {
  const { result } = await ctx.callTool(participantId, 'list_conversations', {
    participantId,
    status: 'active',
  });
  if (result.status !== 'success') throw new ConversationToolError(result);

  const raw = (result.data as { conversations?: unknown } | undefined)?.conversations;
  const list = Array.isArray(raw) ? raw : [];
  return list
    .filter((entry): entry is ConversationMetaLike => typeof entry === 'object' && entry !== null)
    .map((meta) => {
      const id = text(meta.id);
      if (!id) return undefined;
      return {
        id,
        title: text(meta.resolvedTitle) ?? text(meta.sharedTitle) ?? text(meta.title) ?? id,
        messageCount: typeof meta.messageCount === 'number' ? meta.messageCount : 0,
        updatedAt: text(meta.updatedAt) ?? '',
        status: meta.status === 'archived' ? 'archived' : 'active',
      } satisfies ConversationEntry;
    })
    .filter((entry): entry is ConversationEntry => entry !== undefined)
    .sort((a, b) => (a.updatedAt < b.updatedAt ? 1 : a.updatedAt > b.updatedAt ? -1 : 0));
}

export function bindConversation(
  state: TeamsState,
  key: string,
  conversationId: string,
): void {
  const previous = state.legionConversations[key];
  if (previous === conversationId) return;
  if (previous && state.legionConversationTargets[previous] === key) {
    delete state.legionConversationTargets[previous];
  }
  state.legionConversations[key] = conversationId;
  state.legionConversationTargets[conversationId] = key;
}

export function conversationErrorText(error: unknown): string {
  if (error instanceof ConversationToolError) {
    if (error.result.status === 'error') return error.result.error ?? 'unknown error';
    const formatted = formatToolResult(error.result);
    if (formatted) return formatted;
  }
  return error instanceof Error ? error.message : String(error);
}
```

- [ ] **Step 4: Run the tests and typecheck**

Run: `npx vitest run src/conversations.test.ts && npm run typecheck`
Expected: PASS, no type errors.

- [ ] **Step 5: Commit**

```bash
git add src/conversations.ts src/conversations.test.ts
git commit -m "feat(teams): list and bind Legion conversations"
```

---

### Task 3: Picker and result cards

**Files:**
- Modify: `src/conversations.ts`
- Test: `src/conversations.test.ts`

**Interfaces:**
- Consumes: `ConversationEntry`
- Produces: `CONVERSATION_PAGE_SIZE = 5`
- Produces: `conversationCard(entries, currentId, offset, key): IAdaptiveCard`
- Produces: `switchedCard(entry, alreadyCurrent?): IAdaptiveCard`
- Produces: `conversationErrorCard(message): IAdaptiveCard`

- [ ] **Step 1: Write the failing tests**

Append to `src/conversations.test.ts`:

```ts
import type { IAdaptiveCard } from '@microsoft/teams.cards';
import {
  CONVERSATION_PAGE_SIZE,
  conversationCard,
  conversationErrorCard,
  switchedCard,
  type ConversationEntry,
} from './conversations.js';

function entry(id: string, title: string, updatedAt = '2026-11-25T00:00:00.000Z'): ConversationEntry {
  return { id, title, messageCount: 3, updatedAt, status: 'active' };
}

describe('conversationCard', () => {
  it('shows five conversations and a More button', () => {
    const entries = Array.from({ length: 7 }, (_, index) => entry(`c${index}`, `Conversation ${index}`));
    const card = conversationCard(entries, 'c1', 0, 'app:alice') as unknown as {
      body: Array<{ text?: string }>;
      actions: Array<{ title: string; data: Record<string, unknown> }>;
    };
    expect(CONVERSATION_PAGE_SIZE).toBe(5);
    expect(card.actions).toHaveLength(6);
    expect(card.actions.slice(0, 5).map((action) => action.data.conversationId))
      .toEqual(['c0', 'c1', 'c2', 'c3', 'c4']);
    expect(card.actions[5]).toEqual({
      type: 'Action.Execute',
      title: 'More…',
      data: { action: 'switch', offset: 5, key: 'app:alice' },
    });
    expect(card.body.some((block) => block.text === 'Showing 1–5 of 7')).toBe(true);
    expect(JSON.stringify(card)).toContain('· current');
  });

  it('shows the final page without More', () => {
    const entries = Array.from({ length: 7 }, (_, index) => entry(`c${index}`, `Conversation ${index}`));
    const card = conversationCard(entries, undefined, 5, 'app:alice') as unknown as {
      body: Array<{ text?: string }>;
      actions: unknown[];
    };
    expect(card.actions).toHaveLength(2);
    expect(card.body.some((block) => block.text === 'Showing 6–7 of 7')).toBe(true);
  });

  it('marks a current conversation outside the list', () => {
    const card = conversationCard([entry('c0', 'Only')], 'gone', 0, 'app:alice') as unknown as {
      body: Array<{ text?: string }>;
    };
    expect(card.body.some((block) => block.text === 'Current: gone (not in the active list)'))
      .toBe(true);
  });

  it('renders an empty state with no actions', () => {
    const card = conversationCard([], undefined, 0, 'app:alice') as unknown as {
      body: Array<{ text?: string }>;
      actions?: unknown[];
    };
    expect(card.body.some((block) => block.text === '🔀 Switch conversation')).toBe(true);
    expect(card.body.some((block) => block.text === 'No conversations yet. Send a message to start one.'))
      .toBe(true);
    expect(card.actions).toBeUndefined();
  });

  it('truncates long titles and stays under the default card limit', () => {
    const entries = Array.from({ length: 5 }, (_, index) =>
      entry(`c${index}`, `${index}`.padEnd(80, 'x')));
    const card = conversationCard(entries, undefined, 0, 'app:alice');
    const serialized = JSON.stringify(card);
    expect(serialized.length).toBeLessThan(28000);
    expect(serialized).toContain('…');
    expect(serialized).not.toContain('x'.repeat(60));
  });
});

describe('switchedCard', () => {
  it('shows the new binding', () => {
    const card = switchedCard(entry('c1', 'Project X')) as unknown as {
      body: Array<{ text?: string }>;
    };
    expect(card.body.some((block) => block.text === '✅ Switched conversation')).toBe(true);
    expect(JSON.stringify(card)).toContain('Project X');
    expect(JSON.stringify(card)).toContain('Future messages in this chat go to this conversation.');
  });

  it('shows the already-current state', () => {
    const card = switchedCard(entry('c1', 'Project X'), true) as unknown as {
      body: Array<{ text?: string }>;
    };
    expect(card.body.some((block) => block.text === '✅ Already on this conversation')).toBe(true);
  });
});

describe('conversationErrorCard', () => {
  it('shows a warning with no actions', () => {
    const card = conversationErrorCard('Conversation not found.') as unknown as {
      body: Array<{ text?: string }>;
      actions?: unknown[];
    };
    expect(card.body[0]?.text).toBe('⚠️ Conversation not found.');
    expect(card.actions).toBeUndefined();
  });
});

it('keeps the card typed as IAdaptiveCard', () => {
  const card: IAdaptiveCard = conversationCard([], undefined, 0, 'app:alice');
  expect(card.type).toBe('AdaptiveCard');
});
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `npx vitest run src/conversations.test.ts`
Expected: FAIL — the card exports do not exist.

- [ ] **Step 3: Implement the cards in `src/conversations.ts`**

Add the import at the top:

```ts
import type { IAdaptiveCard } from '@microsoft/teams.cards';
```

Append:

```ts
export const CONVERSATION_PAGE_SIZE = 5;

const MAX_TITLE = 40;

function truncate(value: string, max: number): string {
  return value.length <= max ? value : `${value.slice(0, max - 1)}…`;
}

function updatedDate(iso: string): string {
  const date = iso.slice(0, 10);
  return /^\d{4}-\d{2}-\d{2}$/.test(date) ? date : 'unknown';
}

function currentLabel(entries: ConversationEntry[], currentId: string | undefined): string {
  if (!currentId) return 'Current: (none)';
  const current = entries.find((entry) => entry.id === currentId);
  return current
    ? `Current: ${truncate(current.title, MAX_TITLE)}`
    : `Current: ${truncate(currentId, MAX_TITLE)} (not in the active list)`;
}

export function conversationCard(
  entries: ConversationEntry[],
  currentId: string | undefined,
  offset: number,
  key: string,
): IAdaptiveCard {
  const heading = { type: 'TextBlock', text: '🔀 Switch conversation', weight: 'Bolder', wrap: true };
  if (entries.length === 0) {
    return {
      type: 'AdaptiveCard',
      version: '1.4',
      body: [
        heading,
        { type: 'TextBlock', text: 'No conversations yet. Send a message to start one.', wrap: true },
      ],
    } as IAdaptiveCard;
  }

  const start = Math.max(0, Math.min(Math.floor(offset), entries.length - 1));
  const page = entries.slice(start, start + CONVERSATION_PAGE_SIZE);
  const body: unknown[] = [
    heading,
    { type: 'TextBlock', text: currentLabel(entries, currentId), wrap: true },
    {
      type: 'TextBlock',
      text: `Showing ${start + 1}–${start + page.length} of ${entries.length}`,
      wrap: true,
    },
    {
      type: 'FactSet',
      facts: page.map((entry) => ({
        title: truncate(entry.title, MAX_TITLE),
        value: `${entry.messageCount} messages · updated ${updatedDate(entry.updatedAt)}` +
          (entry.id === currentId ? ' · current' : ''),
      })),
    },
  ];
  const actions: unknown[] = page.map((entry) => ({
    type: 'Action.Execute',
    title: truncate(entry.title, MAX_TITLE),
    data: { action: 'switch', conversationId: entry.id, key },
  }));
  if (start + CONVERSATION_PAGE_SIZE < entries.length) {
    actions.push({
      type: 'Action.Execute',
      title: 'More…',
      data: { action: 'switch', offset: start + CONVERSATION_PAGE_SIZE, key },
    });
  }
  return { type: 'AdaptiveCard', version: '1.4', body, actions } as IAdaptiveCard;
}

export function switchedCard(
  entry: ConversationEntry,
  alreadyCurrent = false,
): IAdaptiveCard {
  return {
    type: 'AdaptiveCard',
    version: '1.4',
    body: [
      {
        type: 'TextBlock',
        text: alreadyCurrent ? '✅ Already on this conversation' : '✅ Switched conversation',
        weight: 'Bolder',
        wrap: true,
      },
      {
        type: 'FactSet',
        facts: [
          { title: 'Title', value: truncate(entry.title, MAX_TITLE) },
          { title: 'Messages', value: String(entry.messageCount) },
          { title: 'Updated', value: updatedDate(entry.updatedAt) },
        ],
      },
      {
        type: 'TextBlock',
        text: 'Future messages in this chat go to this conversation.',
        wrap: true,
      },
    ],
  } as IAdaptiveCard;
}

export function conversationErrorCard(message: string): IAdaptiveCard {
  return {
    type: 'AdaptiveCard',
    version: '1.4',
    body: [{ type: 'TextBlock', text: `⚠️ ${message}`, wrap: true }],
  } as IAdaptiveCard;
}
```

- [ ] **Step 4: Run the tests and typecheck**

Run: `npx vitest run src/conversations.test.ts && npm run typecheck`
Expected: PASS, no type errors.

- [ ] **Step 5: Commit**

```bash
git add src/conversations.ts src/conversations.test.ts
git commit -m "feat(teams): build conversation picker cards"
```

---

### Task 4: `resolveChatParticipant` identity helper

**Files:**
- Modify: `src/identity.ts`
- Test: `src/identity.test.ts`

**Interfaces:**
- Produces: `ChatParticipant { participantId: string; key: string }`
- Produces: `resolveChatParticipant(args): ChatParticipant | undefined` with args
  `{ collective, appId, aadObjectId?, conversationId?, isGroup, options }`

- [ ] **Step 1: Write the failing tests**

Append to `src/identity.test.ts`:

```ts
import { resolveChatParticipant } from './identity.js';
import type { CollectiveLike } from './directory.js';
import { resolveOptions } from './options.js';

const participant = (id: string, connector: string, externalId: string, status = 'active') => ({
  id, name: id, type: 'user', status,
  identities: [{ connector, externalId }],
});

function collectiveOf(participants: Array<ReturnType<typeof participant>>): CollectiveLike {
  return {
    list: () => participants,
    get: (id: string) => participants.find((candidate) => candidate.id === id),
    findByIdentity: (connector: string, externalId: string) =>
      participants.find((candidate) =>
        candidate.identities.some((identity) =>
          identity.connector === connector && identity.externalId === externalId)),
  } as unknown as CollectiveLike;
}

describe('resolveChatParticipant', () => {
  const alice = participant('alice', 'teams', 'aad-1');
  const group = participant('team-eng', 'teams-group', '19:g@thread.v2');
  const operator = participant('operator', 'teams', 'aad-op');

  it('resolves a registered personal user', () => {
    expect(resolveChatParticipant({
      collective: collectiveOf([alice, operator]),
      appId: 'app',
      aadObjectId: 'aad-1',
      isGroup: false,
      options: resolveOptions({ stateDir: '/tmp/x' }),
    })).toEqual({ participantId: 'alice', key: 'app:alice' });
  });

  it('resolves a personal fallback with the AAD suffix', () => {
    expect(resolveChatParticipant({
      collective: collectiveOf([alice, operator]),
      appId: 'app',
      aadObjectId: 'aad-9',
      isGroup: false,
      options: resolveOptions({
        stateDir: '/tmp/x',
        unknownUserPolicy: 'fallback',
        fallbackUserId: 'operator',
      }),
    })).toEqual({ participantId: 'operator', key: 'app:operator:aad-9' });
  });

  it('resolves a registered group', () => {
    expect(resolveChatParticipant({
      collective: collectiveOf([alice, group]),
      appId: 'app',
      conversationId: '19:g@thread.v2',
      isGroup: true,
      options: resolveOptions({ stateDir: '/tmp/x' }),
    })).toEqual({ participantId: 'team-eng', key: 'app:team-eng:19:g@thread.v2' });
  });

  it('resolves a group fallback', () => {
    expect(resolveChatParticipant({
      collective: collectiveOf([alice, operator]),
      appId: 'app',
      conversationId: '19:unknown@thread.v2',
      isGroup: true,
      options: resolveOptions({
        stateDir: '/tmp/x',
        unknownUserPolicy: 'fallback',
        fallbackUserId: 'operator',
      }),
    })).toEqual({ participantId: 'operator', key: 'app:operator:19:unknown@thread.v2' });
  });

  it('returns undefined without a fallback', () => {
    expect(resolveChatParticipant({
      collective: collectiveOf([alice]),
      appId: 'app',
      aadObjectId: 'aad-9',
      isGroup: false,
      options: resolveOptions({ stateDir: '/tmp/x' }),
    })).toBeUndefined();
    expect(resolveChatParticipant({
      collective: collectiveOf([alice]),
      appId: 'app',
      conversationId: '19:unknown@thread.v2',
      isGroup: true,
      options: resolveOptions({ stateDir: '/tmp/x' }),
    })).toBeUndefined();
  });

  it('rejects a retired fallback participant', () => {
    expect(resolveChatParticipant({
      collective: collectiveOf([alice, participant('operator', 'teams', 'aad-op', 'retired')]),
      appId: 'app',
      aadObjectId: 'aad-9',
      isGroup: false,
      options: resolveOptions({
        stateDir: '/tmp/x',
        unknownUserPolicy: 'fallback',
        fallbackUserId: 'operator',
      }),
    })).toBeUndefined();
  });
});
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `npx vitest run src/identity.test.ts`
Expected: FAIL — `resolveChatParticipant` is not exported.

- [ ] **Step 3: Implement the helper**

Append to `src/identity.ts`:

```ts
export interface ChatParticipant {
  participantId: string;
  key: string;
}

function fallbackParticipant(
  collective: CollectiveLike,
  options: ResolvedOptions,
) {
  if (options.unknownUserPolicy !== 'fallback' || !options.fallbackUserId) return undefined;
  const fallback = collective.get(options.fallbackUserId);
  return fallback && (fallback.status ?? 'active') !== 'retired' ? fallback : undefined;
}

export function resolveChatParticipant(args: {
  collective: CollectiveLike;
  appId: string;
  aadObjectId?: string;
  conversationId?: string;
  isGroup: boolean;
  options: ResolvedOptions;
}): ChatParticipant | undefined {
  const { collective, appId, aadObjectId, conversationId, isGroup, options } = args;

  if (isGroup) {
    if (!conversationId) return undefined;
    const group = findGroupActor(collective, conversationId);
    if (group) return { participantId: group.id, key: `${appId}:${group.id}:${conversationId}` };
    const fallback = fallbackParticipant(collective, options);
    return fallback
      ? { participantId: fallback.id, key: `${appId}:${fallback.id}:${conversationId}` }
      : undefined;
  }

  const user = aadObjectId ? findTeamsActor(collective, aadObjectId) : undefined;
  if (user) return { participantId: user.id, key: `${appId}:${user.id}` };
  const fallback = fallbackParticipant(collective, options);
  return fallback && aadObjectId
    ? { participantId: fallback.id, key: `${appId}:${fallback.id}:${aadObjectId}` }
    : undefined;
}
```

- [ ] **Step 4: Run the tests and typecheck**

Run: `npx vitest run src/identity.test.ts && npm run typecheck`
Expected: PASS, no type errors.

- [ ] **Step 5: Commit**

```bash
git add src/identity.ts src/identity.test.ts
git commit -m "feat(teams): resolve chat participants for card actions"
```

---

### Task 5: Typing helper and command wiring

**Files:**
- Modify: `src/commands.ts`
- Modify: `src/inbound.ts`
- Modify: `src/commands.test.ts`
- Test: `src/inbound.test.ts`

**Interfaces:**
- Produces: `CommandDeps.sendCard: (card: IAdaptiveCard) => Promise<void>`
- Produces: `CommandDeps.stopTyping?: () => void`
- Produces: `startTyping(context: TeamsContextLike, intervalMs: number): () => void` (module-private in `inbound.ts`)

- [ ] **Step 1: Write the failing tests**

Add to `src/inbound.test.ts` (import `TypingActivityInput` from `@microsoft/teams.api`):

```ts
import { TypingActivityInput } from '@microsoft/teams.api';

it('sends a typing indicator while handling a command', async () => {
  const s = setup([alice]);
  const send = vi.fn(async () => undefined);
  await handleInbound(
    inboundDeps(s, { typingIntervalMs: 60_000 }),
    { activity: activity('/new'), send },
  );
  expect(send).toHaveBeenCalledWith(expect.any(TypingActivityInput));
  expect(s.send).toBeDefined();
});
```

Update the `deps()` helper in `src/commands.test.ts` to include `sendCard`:

```ts
import type { IAdaptiveCard } from '@microsoft/teams.cards';

function deps(text: string, overrides: Record<string, unknown> = {}) {
  const sent: string[] = [];
  const cards: IAdaptiveCard[] = [];
  const callTool = vi.fn(async () => ({ result: { status: 'success' }, conversationId: 'c' }));
  return {
    sent,
    cards,
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
      sendCard: async (card: IAdaptiveCard) => { cards.push(card); },
      ...overrides,
    },
  };
}
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `npx vitest run src/commands.test.ts src/inbound.test.ts`
Expected: FAIL — `sendCard` is required by the type after Step 3, and the typing test fails before `startTyping` exists. (Vitest transpiles without type checking; the typing assertion is the observable failure.)

- [ ] **Step 3: Implement**

In `src/commands.ts`, add the import and fields:

```ts
import type { IAdaptiveCard } from '@microsoft/teams.cards';
```

```ts
export interface CommandDeps {
  // ...existing fields...
  send: (text: string) => Promise<void>;
  sendCard: (card: IAdaptiveCard) => Promise<void>;
  stopTyping?: () => void;
}
```

In `src/inbound.ts`, add the import:

```ts
import type { IAdaptiveCard } from '@microsoft/teams.cards';
```

Add `startTyping` before `runTurn`:

```ts
function startTyping(context: TeamsContextLike, intervalMs: number): () => void {
  if (intervalMs <= 0) return () => undefined;
  const tick = () => context.send(new TypingActivityInput()).catch(() => undefined);
  tick();
  const interval = setInterval(tick, intervalMs);
  let stopped = false;
  return () => {
    if (stopped) return;
    stopped = true;
    clearInterval(interval);
  };
}
```

Replace the inline typing block in `runTurn` with:

```ts
const stopTyping = startTyping(context, deps.options.typingIntervalMs);
```

Keep the existing `try { ... } finally { stopTyping(); }` structure.

In `handleInbound`, after `reply`, add:

```ts
const sendCard = async (card: IAdaptiveCard) => {
  const activity = new MessageActivityInput().addCard('adaptive', card);
  if (deps.options.aiLabel) activity.addAiGenerated();
  await context.send(activity);
};
```

Wrap both command paths. Group path:

```ts
const stopTyping = startTyping(context, deps.options.typingIntervalMs);
try {
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
    sendCard,
    stopTyping,
  });
  if (handled) {
    await deps.stateStore.save(deps.state);
    return;
  }
} finally {
  stopTyping();
}
```

Personal path:

```ts
const stopTyping = startTyping(context, deps.options.typingIntervalMs);
try {
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
    sendCard,
    stopTyping,
  });
  if (handled) {
    await deps.stateStore.save(deps.state);
    return;
  }
} finally {
  stopTyping();
}
```

- [ ] **Step 4: Run the tests and typecheck**

Run: `npx vitest run src/commands.test.ts src/inbound.test.ts && npm run typecheck`
Expected: PASS, no type errors.

- [ ] **Step 5: Commit**

```bash
git add src/commands.ts src/inbound.ts src/commands.test.ts src/inbound.test.ts
git commit -m "feat(teams): wire card sending and typing into commands"
```

---

### Task 6: `/switch` command

**Files:**
- Modify: `src/commands.ts`
- Modify: `src/commands.test.ts`
- Modify: `src/inbound.test.ts`

**Interfaces:**
- Consumes: `listConversations`, `bindConversation`, `conversationCard`, `conversationErrorText`
- Consumes: `CommandDeps.sendCard`

- [ ] **Step 1: Write the failing tests**

Add to `src/commands.test.ts`:

```ts
it('sends the conversation picker card on /switch', async () => {
  const { deps: d, cards } = deps('/switch');
  const callTool = vi.fn(async () => ({
    result: { status: 'success', data: { conversations: [
      { id: 'legion-1', resolvedTitle: 'First', messageCount: 3,
        updatedAt: '2026-11-24T00:00:00.000Z', status: 'active' },
    ] } },
    conversationId: 'c',
  }));
  await handleCommand({ ...d, ctx: { callTool } as unknown as ConnectorContext });
  expect(callTool).toHaveBeenCalledWith('alice', 'list_conversations', {
    participantId: 'alice',
    status: 'active',
  });
  expect(cards).toHaveLength(1);
  expect(JSON.stringify(cards[0])).toContain('First');
});

it('switches directly to a conversation id and cleans the old target', async () => {
  const { deps: d, sent } = deps('/switch legion-2');
  d.state.legionConversations['app:alice'] = 'legion-1';
  d.state.legionConversationTargets['legion-1'] = 'app:alice';
  const callTool = vi.fn(async () => ({
    result: { status: 'success', data: { conversations: [
      { id: 'legion-2', resolvedTitle: 'Second', messageCount: 2,
        updatedAt: '2026-11-25T00:00:00.000Z', status: 'active' },
    ] } },
    conversationId: 'c',
  }));
  await handleCommand({ ...d, ctx: { callTool } as unknown as ConnectorContext });
  expect(d.state.legionConversations['app:alice']).toBe('legion-2');
  expect(d.state.legionConversationTargets['legion-1']).toBeUndefined();
  expect(d.state.legionConversationTargets['legion-2']).toBe('app:alice');
  expect(sent[0]).toBe('Switched to Second (2 messages).');
});

it('reports the current conversation on a direct switch', async () => {
  const { deps: d, sent } = deps('/switch legion-1');
  d.state.legionConversations['app:alice'] = 'legion-1';
  const callTool = vi.fn(async () => ({
    result: { status: 'success', data: { conversations: [
      { id: 'legion-1', resolvedTitle: 'First', messageCount: 3,
        updatedAt: '2026-11-24T00:00:00.000Z', status: 'active' },
    ] } },
    conversationId: 'c',
  }));
  await handleCommand({ ...d, ctx: { callTool } as unknown as ConnectorContext });
  expect(sent[0]).toBe('Already on First.');
});

it('reports an unknown conversation id', async () => {
  const { deps: d, sent } = deps('/switch missing');
  const callTool = vi.fn(async () => ({
    result: { status: 'success', data: { conversations: [] } },
    conversationId: 'c',
  }));
  await handleCommand({ ...d, ctx: { callTool } as unknown as ConnectorContext });
  expect(sent[0]).toBe('Conversation not found. Use /switch to list active conversations.');
});

it('reports a listing failure', async () => {
  const { deps: d, sent } = deps('/switch');
  const callTool = vi.fn(async () => ({
    result: { status: 'error', error: 'not authorized' },
    conversationId: 'c',
  }));
  await handleCommand({ ...d, ctx: { callTool } as unknown as ConnectorContext });
  expect(sent[0]).toBe('Unable to list conversations: not authorized');
});

it('uses the group key for /switch in a group chat', async () => {
  const { deps: d, sent } = deps('/switch legion-2', {
    participantId: 'team-eng',
    isGroup: true,
    conversationKey: 'app:team-eng:19:g@thread.v2',
    conversationId: '19:g@thread.v2',
  });
  d.state.legionConversations['app:team-eng:19:g@thread.v2'] = 'legion-1';
  d.state.legionConversationTargets['legion-1'] = 'app:team-eng:19:g@thread.v2';
  const callTool = vi.fn(async () => ({
    result: { status: 'success', data: { conversations: [
      { id: 'legion-2', resolvedTitle: 'Second', messageCount: 2,
        updatedAt: '2026-11-25T00:00:00.000Z', status: 'active' },
    ] } },
    conversationId: 'c',
  }));
  await handleCommand({ ...d, ctx: { callTool } as unknown as ConnectorContext });
  expect(callTool).toHaveBeenCalledWith('team-eng', 'list_conversations', expect.anything());
  expect(d.state.legionConversations['app:team-eng:19:g@thread.v2']).toBe('legion-2');
  expect(d.state.legionConversationTargets['legion-1']).toBeUndefined();
  expect(d.state.legionConversationTargets['legion-2']).toBe('app:team-eng:19:g@thread.v2');
  expect(sent[0]).toContain('Switched to Second');
});

it('uses the fallback suffix key for /switch in a personal chat', async () => {
  const { deps: d, sent } = deps('/switch legion-2', {
    participantId: 'operator',
    isFallback: true,
    conversationKey: 'app:operator:aad-9',
  });
  d.state.legionConversations['app:operator:aad-9'] = 'legion-1';
  d.state.legionConversationTargets['legion-1'] = 'app:operator:aad-9';
  const callTool = vi.fn(async () => ({
    result: { status: 'success', data: { conversations: [
      { id: 'legion-2', resolvedTitle: 'Second', messageCount: 2,
        updatedAt: '2026-11-25T00:00:00.000Z', status: 'active' },
    ] } },
    conversationId: 'c',
  }));
  await handleCommand({ ...d, ctx: { callTool } as unknown as ConnectorContext });
  expect(callTool).toHaveBeenCalledWith('operator', 'list_conversations', expect.anything());
  expect(d.state.legionConversations['app:operator:aad-9']).toBe('legion-2');
  expect(d.state.legionConversationTargets['legion-1']).toBeUndefined();
  expect(d.state.legionConversationTargets['legion-2']).toBe('app:operator:aad-9');
  expect(sent[0]).toContain('Switched to Second');
});
```

Add to `src/inbound.test.ts`:

```ts
it('sends the conversation picker card for /switch', async () => {
  const s = setup([alice]);
  const send = vi.fn(async () => undefined);
  await handleInbound(inboundDeps(s), { activity: activity('/switch'), send });
  expect(send).toHaveBeenCalledWith(expect.objectContaining({
    attachments: [expect.objectContaining({
      contentType: 'application/vnd.microsoft.card.adaptive',
    })],
  }));
});
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `npx vitest run src/commands.test.ts src/inbound.test.ts -t "switch"`
Expected: FAIL — `/switch` is treated as plain text (`handleCommand` returns `false`).

- [ ] **Step 3: Implement the `/switch` case**

In `src/commands.ts`, add imports:

```ts
import {
  bindConversation,
  conversationCard,
  conversationErrorText,
  listConversations,
} from './conversations.js';
```

Add the case before `default`:

```ts
case '/switch': {
  const currentId = deps.state.legionConversations[key];
  try {
    const entries = await listConversations(deps.ctx, deps.participantId);
    const target = rest[0];
    if (!target) {
      await deps.sendCard(conversationCard(entries, currentId, 0, key));
      return true;
    }
    const entry = entries.find((candidate) => candidate.id === target);
    if (!entry) {
      await deps.send('Conversation not found. Use /switch to list active conversations.');
      return true;
    }
    if (entry.id === currentId) {
      await deps.send(`Already on ${entry.title}.`);
      return true;
    }
    bindConversation(deps.state, key, entry.id);
    await deps.send(`Switched to ${entry.title} (${entry.messageCount} messages).`);
    return true;
  } catch (error) {
    await deps.send(`Unable to list conversations: ${conversationErrorText(error)}`);
    return true;
  }
}
```

- [ ] **Step 4: Run the tests and typecheck**

Run: `npx vitest run src/commands.test.ts src/inbound.test.ts && npm run typecheck`
Expected: PASS, no type errors.

- [ ] **Step 5: Commit**

```bash
git add src/commands.ts src/commands.test.ts src/inbound.test.ts
git commit -m "feat(teams): add /switch with conversation picker"
```

---

### Task 7: `/compact` command

**Files:**
- Modify: `src/commands.ts`
- Modify: `src/commands.test.ts`

**Interfaces:**
- Consumes: `ResolvedOptions.compaction`
- Consumes: `formatToolResult`

- [ ] **Step 1: Write the failing tests**

Add to `src/commands.test.ts`:

```ts
it('reports missing compaction configuration without calling a tool', async () => {
  const { deps: d, sent, callTool } = deps('/compact');
  await handleCommand(d);
  expect(callTool).not.toHaveBeenCalled();
  expect(sent[0]).toBe(
    'Compaction is not configured. Set the "compaction.agentId" connector option.',
  );
});

it('reports no conversation', async () => {
  const { deps: d, sent } = deps('/compact', {
    options: resolveOptions({ stateDir: '/tmp/x', compaction: { agentId: 'summarizer' } }),
  });
  await handleCommand(d);
  expect(sent[0]).toBe('Nothing to compact — no conversation yet.');
});

it('compacts the active chain and reports the summary', async () => {
  const { deps: d, sent } = deps('/compact be brief', {
    options: resolveOptions({
      stateDir: '/tmp/x',
      compaction: { agentId: 'summarizer', instruction: 'Default instruction.' },
    }),
  });
  d.state.legionConversations['app:alice'] = 'legion-1';
  let getCalls = 0;
  const callTool = vi.fn(async (_participant: string, tool: string) => {
    if (tool === 'get_conversation') {
      getCalls += 1;
      return {
        result: {
          status: 'success',
          data: {
            messages: getCalls === 1
              ? [{ id: 'm1', content: 'one' }, { id: 'm2', content: 'two' }]
              : [{ id: 's1', content: 'Summary text' }],
          },
        },
        conversationId: 'c',
      };
    }
    return {
      result: { status: 'success', data: { summaryMessageId: 's1' } },
      conversationId: 'c',
    };
  });
  const stopTyping = vi.fn();
  await handleCommand({
    ...d,
    ctx: { callTool } as unknown as ConnectorContext,
    stopTyping,
  });
  expect(callTool).toHaveBeenCalledWith('alice', 'compact_conversation', {
    conversationId: 'legion-1',
    messageIds: ['m1', 'm2'],
    agentId: 'summarizer',
    instruction: 'be brief',
  });
  expect(stopTyping).toHaveBeenCalledOnce();
  expect(sent[0]).toBe('🗜 Compacted 2 messages into a summary:\n\nSummary text');
});

it('uses the default instruction for a bare /compact', async () => {
  const { deps: d } = deps('/compact', {
    options: resolveOptions({
      stateDir: '/tmp/x',
      compaction: { agentId: 'summarizer', instruction: 'Default instruction.' },
    }),
  });
  d.state.legionConversations['app:alice'] = 'legion-1';
  const callTool = vi.fn(async (_participant: string, tool: string) => {
    if (tool === 'get_conversation') {
      return { result: { status: 'success', data: { messages: [{ id: 'm1' }] } }, conversationId: 'c' };
    }
    return { result: { status: 'success', data: { summaryMessageId: 's1' } }, conversationId: 'c' };
  });
  await handleCommand({ ...d, ctx: { callTool } as unknown as ConnectorContext });
  expect(callTool).toHaveBeenCalledWith(
    'alice',
    'compact_conversation',
    expect.objectContaining({ instruction: 'Default instruction.' }),
  );
});

it('reports an empty conversation without compacting', async () => {
  const { deps: d, sent } = deps('/compact', {
    options: resolveOptions({ stateDir: '/tmp/x', compaction: { agentId: 'summarizer' } }),
  });
  d.state.legionConversations['app:alice'] = 'legion-1';
  const callTool = vi.fn(async () => ({
    result: { status: 'success', data: { messages: [] } },
    conversationId: 'c',
  }));
  await handleCommand({ ...d, ctx: { callTool } as unknown as ConnectorContext });
  expect(sent[0]).toBe('Nothing to compact — conversation is empty.');
  expect(callTool).toHaveBeenCalledTimes(1);
});

it('reports a compaction tool error', async () => {
  const { deps: d, sent } = deps('/compact', {
    options: resolveOptions({ stateDir: '/tmp/x', compaction: { agentId: 'summarizer' } }),
  });
  d.state.legionConversations['app:alice'] = 'legion-1';
  const callTool = vi.fn(async (_participant: string, tool: string) => {
    if (tool === 'get_conversation') {
      return { result: { status: 'success', data: { messages: [{ id: 'm1' }] } }, conversationId: 'c' };
    }
    return { result: { status: 'error', error: 'boom' }, conversationId: 'c' };
  });
  await handleCommand({ ...d, ctx: { callTool } as unknown as ConnectorContext });
  expect(sent[0]).toBe('Compaction failed: boom');
});

it('reports a pending compaction approval', async () => {
  const { deps: d, sent } = deps('/compact', {
    options: resolveOptions({ stateDir: '/tmp/x', compaction: { agentId: 'summarizer' } }),
  });
  d.state.legionConversations['app:alice'] = 'legion-1';
  const callTool = vi.fn(async (_participant: string, tool: string) => {
    if (tool === 'get_conversation') {
      return { result: { status: 'success', data: { messages: [{ id: 'm1' }] } }, conversationId: 'c' };
    }
    return { result: { status: 'pending_approval' }, conversationId: 'c' };
  });
  await handleCommand({ ...d, ctx: { callTool } as unknown as ConnectorContext });
  expect(sent[0]).toBe('Compaction failed: Awaiting approval.');
});
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `npx vitest run src/commands.test.ts -t "compact"`
Expected: FAIL — `/compact` returns `false` and sends nothing.

- [ ] **Step 3: Implement the `/compact` case**

In `src/commands.ts`, add the import:

```ts
import { formatToolResult } from './render.js';
```

Add the helper above `handleCommand`:

```ts
function toolFailure(result: { status: string; error?: string; message?: string }): string {
  if (result.status === 'error') return result.error ?? 'unknown error';
  return formatToolResult(result) || result.status;
}
```

Add the case before `default`:

```ts
case '/compact': {
  const compaction = deps.options.compaction;
  if (!compaction) {
    await deps.send('Compaction is not configured. Set the "compaction.agentId" connector option.');
    return true;
  }
  const conversationId = deps.state.legionConversations[key];
  if (!conversationId) {
    await deps.send('Nothing to compact — no conversation yet.');
    return true;
  }
  try {
    const loaded = await deps.ctx.callTool(deps.participantId, 'get_conversation', {
      conversationId,
    });
    if (loaded.result.status !== 'success') {
      await deps.send(`Compaction failed: ${toolFailure(loaded.result)}`);
      return true;
    }
    const messages = (loaded.result.data as { messages?: Array<{ id?: unknown }> } | undefined)
      ?.messages ?? [];
    const messageIds = messages
      .map((message) => message.id)
      .filter((id): id is string => typeof id === 'string' && id.length > 0);
    if (messageIds.length === 0) {
      await deps.send('Nothing to compact — conversation is empty.');
      return true;
    }
    const instruction = rest.join(' ').trim() || compaction.instruction;
    const compacted = await deps.ctx.callTool(deps.participantId, 'compact_conversation', {
      conversationId,
      messageIds,
      agentId: compaction.agentId,
      ...(instruction ? { instruction } : {}),
    });
    if (compacted.result.status !== 'success') {
      await deps.send(`Compaction failed: ${toolFailure(compacted.result)}`);
      return true;
    }
    const summaryMessageId = (compacted.result.data as { summaryMessageId?: unknown } | undefined)
      ?.summaryMessageId;
    if (typeof summaryMessageId !== 'string' || !summaryMessageId) {
      await deps.send(`Compacted ${messageIds.length} messages — no summary message was returned.`);
      return true;
    }
    const refreshed = await deps.ctx.callTool(deps.participantId, 'get_conversation', {
      conversationId,
    });
    const summary = refreshed.result.status === 'success'
      ? ((refreshed.result.data as {
        messages?: Array<{ id?: unknown; content?: unknown }>;
      } | undefined)?.messages ?? []).find((message) => message.id === summaryMessageId)
      : undefined;
    const content = typeof summary?.content === 'string' ? summary.content.trim() : '';
    if (!content) {
      await deps.send(`Compacted ${messageIds.length} messages — summary unavailable.`);
      return true;
    }
    deps.stopTyping?.();
    await deps.send(
      `🗜 Compacted ${messageIds.length} messages into a summary:\n\n${content.slice(0, 3500)}`,
    );
    return true;
  } catch (error) {
    await deps.send(`Compaction failed: ${error instanceof Error ? error.message : String(error)}`);
    return true;
  }
}
```

- [ ] **Step 4: Run the tests and typecheck**

Run: `npx vitest run src/commands.test.ts && npm run typecheck`
Expected: PASS, no type errors.

- [ ] **Step 5: Commit**

```bash
git add src/commands.ts src/commands.test.ts
git commit -m "feat(teams): add /compact command"
```

---

### Task 8: `/context` command

**Files:**
- Modify: `src/commands.ts`
- Modify: `src/commands.test.ts`

**Interfaces:**
- Consumes: `listConversations`, `toolFailure`
- Produces: local `formatUsage(totals)` and `formatCount(value)` helpers

- [ ] **Step 1: Write the failing tests**

Add to `src/commands.test.ts`:

```ts
it('shows conversation and usage details on /context', async () => {
  const { deps: d, sent } = deps('/context');
  d.state.legionConversations['app:alice'] = 'legion-1';
  const callTool = vi.fn(async (_participant: string, tool: string) => {
    if (tool === 'get_conversation') {
      return {
        result: {
          status: 'success',
          data: {
            title: 'Fallback title',
            status: 'active',
            messages: [{ id: 'm1' }, { id: 'm2' }],
          },
        },
        conversationId: 'c',
      };
    }
    if (tool === 'list_conversations') {
      return {
        result: {
          status: 'success',
          data: {
            conversations: [{
              id: 'legion-1', resolvedTitle: 'Project X', messageCount: 10,
              updatedAt: '2026-11-25T00:00:00.000Z', status: 'active',
            }],
          },
        },
        conversationId: 'c',
      };
    }
    return {
      result: {
        status: 'success',
        data: {
          totals: {
            messageCount: 10, input: 1234, output: 567, reasoning: 89,
            cache: { read: 100, write: 50 }, cost: 0.1234,
          },
        },
      },
      conversationId: 'c',
    };
  });
  await handleCommand({ ...d, ctx: { callTool } as unknown as ConnectorContext });
  expect(callTool).toHaveBeenCalledWith('alice', 'query_usage', { conversationId: 'legion-1' });
  expect(sent[0]).toContain('📊 Context — Project X');
  expect(sent[0]).toContain('Conversation: legion-1');
  expect(sent[0]).toContain('Status: active · Messages: 10 total, 2 active');
  expect(sent[0]).toContain('Tokens: in 1,234 · out 567 · reasoning 89');
  expect(sent[0]).toContain('Cache: 100 read / 50 write · Cost: $0.1234');
});

it('degrades to active count when the conversation meta is unavailable', async () => {
  const { deps: d, sent } = deps('/context');
  d.state.legionConversations['app:alice'] = 'legion-1';
  const callTool = vi.fn(async (_participant: string, tool: string) => {
    if (tool === 'get_conversation') {
      return {
        result: { status: 'success', data: { title: 'Fallback title', status: 'active', messages: [{ id: 'm1' }] } },
        conversationId: 'c',
      };
    }
    if (tool === 'list_conversations') {
      return { result: { status: 'error', error: 'unavailable' }, conversationId: 'c' };
    }
    return { result: { status: 'error', error: 'usage offline' }, conversationId: 'c' };
  });
  await handleCommand({ ...d, ctx: { callTool } as unknown as ConnectorContext });
  expect(sent[0]).toContain('📊 Context — Fallback title');
  expect(sent[0]).toContain('Messages: 1 active');
  expect(sent[0]).toContain('Usage: unavailable (usage offline)');
});

it('reports no usage recorded', async () => {
  const { deps: d, sent } = deps('/context');
  d.state.legionConversations['app:alice'] = 'legion-1';
  const callTool = vi.fn(async (_participant: string, tool: string) => {
    if (tool === 'get_conversation') {
      return { result: { status: 'success', data: { title: 'Thread', status: 'active', messages: [] } }, conversationId: 'c' };
    }
    if (tool === 'list_conversations') {
      return { result: { status: 'success', data: { conversations: [] } }, conversationId: 'c' };
    }
    return { result: { status: 'success', data: {} }, conversationId: 'c' };
  });
  await handleCommand({ ...d, ctx: { callTool } as unknown as ConnectorContext });
  expect(sent[0]).toContain('Usage: no usage recorded yet.');
});

it('reports no conversation on /context', async () => {
  const { deps: d, sent } = deps('/context');
  await handleCommand(d);
  expect(sent[0]).toBe('No conversation yet — send a message first.');
});

it('reports an unavailable conversation on /context', async () => {
  const { deps: d, sent } = deps('/context');
  d.state.legionConversations['app:alice'] = 'legion-1';
  const callTool = vi.fn(async () => ({
    result: { status: 'error', error: 'not found' },
    conversationId: 'c',
  }));
  await handleCommand({ ...d, ctx: { callTool } as unknown as ConnectorContext });
  expect(sent[0]).toBe('Context unavailable: not found');
});
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `npx vitest run src/commands.test.ts -t "context"`
Expected: FAIL — `/context` returns `false`.

- [ ] **Step 3: Implement the `/context` case**

In `src/commands.ts`, add the type import:

```ts
import type { ConversationEntry } from './conversations.js';
```

Add helpers above `handleCommand`:

```ts
function formatCount(value: unknown): string {
  const number = typeof value === 'number' && Number.isFinite(value) ? value : 0;
  return number.toLocaleString('en-US');
}

function formatUsage(totals: {
  input?: unknown;
  output?: unknown;
  reasoning?: unknown;
  cache?: { read?: unknown; write?: unknown };
  cost?: unknown;
}): string {
  const cost = typeof totals.cost === 'number' && Number.isFinite(totals.cost) ? totals.cost : 0;
  return [
    `Tokens: in ${formatCount(totals.input)} · out ${formatCount(totals.output)}` +
      ` · reasoning ${formatCount(totals.reasoning)}`,
    `Cache: ${formatCount(totals.cache?.read)} read / ${formatCount(totals.cache?.write)} write` +
      ` · Cost: $${cost.toFixed(4)}`,
  ].join('\n');
}
```

Add the case before `default`:

```ts
case '/context': {
  const conversationId = deps.state.legionConversations[key];
  if (!conversationId) {
    await deps.send('No conversation yet — send a message first.');
    return true;
  }
  let loaded;
  try {
    loaded = await deps.ctx.callTool(deps.participantId, 'get_conversation', {
      conversationId,
    });
  } catch (error) {
    await deps.send(`Context unavailable: ${error instanceof Error ? error.message : String(error)}`);
    return true;
  }
  if (loaded.result.status !== 'success') {
    await deps.send(`Context unavailable: ${toolFailure(loaded.result)}`);
    return true;
  }
  const data = loaded.result.data as {
    title?: unknown;
    status?: unknown;
    messages?: unknown;
  } | undefined;
  const activeMessages = Array.isArray(data?.messages) ? data.messages.length : 0;
  const fallbackTitle =
    typeof data?.title === 'string' && data.title.trim() ? data.title.trim() : conversationId;
  const status =
    typeof data?.status === 'string' && data.status.trim() ? data.status.trim() : 'unknown';

  let meta: ConversationEntry | undefined;
  try {
    const entries = await listConversations(deps.ctx, deps.participantId);
    meta = entries.find((entry) => entry.id === conversationId);
  } catch {
    meta = undefined;
  }

  let usageText: string;
  try {
    const usage = await deps.ctx.callTool(deps.participantId, 'query_usage', {
      conversationId,
    });
    if (usage.result.status !== 'success') {
      usageText = `Usage: unavailable (${toolFailure(usage.result)})`;
    } else {
      const totals = (usage.result.data as { totals?: Parameters<typeof formatUsage>[0] } | undefined)
        ?.totals;
      usageText = totals ? formatUsage(totals) : 'Usage: no usage recorded yet.';
    }
  } catch (error) {
    usageText = `Usage: unavailable (${error instanceof Error ? error.message : String(error)})`;
  }

  const title = meta?.title ?? fallbackTitle;
  const counts = meta
    ? `Messages: ${meta.messageCount} total, ${activeMessages} active`
    : `Messages: ${activeMessages} active`;
  await deps.send([
    `📊 Context — ${title}`,
    `Conversation: ${conversationId}`,
    `Status: ${status} · ${counts}`,
    usageText,
  ].join('\n'));
  return true;
}
```

- [ ] **Step 4: Run the tests and typecheck**

Run: `npx vitest run src/commands.test.ts && npm run typecheck`
Expected: PASS, no type errors.

- [ ] **Step 5: Commit**

```bash
git add src/commands.ts src/commands.test.ts
git commit -m "feat(teams): add /context command"
```

---

### Task 9: Help text

**Files:**
- Modify: `src/commands.ts`
- Modify: `src/commands.test.ts`

**Interfaces:**
- No new interfaces; updates the `/help` copy.

- [ ] **Step 1: Write the failing tests**

Replace the existing `/help` test in `src/commands.test.ts`:

```ts
it('lists the new commands and group scoping in /help', async () => {
  const { deps: d, sent } = deps('/help');
  await expect(handleCommand(d)).resolves.toBe(true);
  expect(sent[0]).toContain('/switch — list and switch to an existing conversation');
  expect(sent[0]).toContain('/switch <conversation-id> — switch directly');
  expect(sent[0]).toContain('/compact [instruction] — summarize/compact this conversation');
  expect(sent[0]).toContain('/context — show usage and context details');
  expect(sent[0]).toContain(
    'In a group chat, /new, /switch, /compact, and /context act on the group; /pair runs as you.',
  );
});
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `npx vitest run src/commands.test.ts -t "help"`
Expected: FAIL — the help text lacks the new commands.

- [ ] **Step 3: Update the `/help` text**

Replace the `/help` case body in `src/commands.ts` with:

```ts
case '/help': {
  await deps.send(
    'Commands:\n' +
    '/new — start a new conversation\n' +
    '/switch — list and switch to an existing conversation\n' +
    '/switch <conversation-id> — switch directly\n' +
    '/compact [instruction] — summarize/compact this conversation\n' +
    '/context — show usage and context details\n' +
    '/pair <code> <participant-id> — pair a pending Teams identity\n' +
    '/whoami — show your identity\n' +
    '/help — this message\n' +
    'In a group chat, /new, /switch, /compact, and /context act on the group; /pair runs as you.',
  );
  return true;
}
```

- [ ] **Step 4: Run the tests and typecheck**

Run: `npx vitest run src/commands.test.ts && npm run typecheck`
Expected: PASS, no type errors.

- [ ] **Step 5: Commit**

```bash
git add src/commands.ts src/commands.test.ts
git commit -m "docs(teams): list conversation commands in help"
```

---

### Task 10: Auto-provision command tool grants

**Files:**
- Modify: `src/inbound.ts`
- Modify: `src/inbound.test.ts`

**Interfaces:**
- No new interfaces; extends the `tools` map used by `autoProvision` and `autoProvisionGroup`.

- [ ] **Step 1: Update the failing test expectations**

In `src/inbound.test.ts`, update the three exact `tools` expectations:

```ts
tools: {
  communicate: 'auto',
  approval_response: 'auto',
  list_conversations: 'auto',
  get_conversation: 'auto',
  query_usage: 'auto',
  compact_conversation: 'auto',
},
```

At the `modify_user` assertion, use:

```ts
tools: {
  search: 'auto',
  communicate: 'auto',
  approval_response: 'auto',
  list_conversations: 'auto',
  get_conversation: 'auto',
  query_usage: 'auto',
  compact_conversation: 'auto',
},
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `npx vitest run src/inbound.test.ts`
Expected: FAIL — the provisioned tool map lacks the four new grants.

- [ ] **Step 3: Implement the grants**

In both `autoProvision` and `autoProvisionGroup` in `src/inbound.ts`, replace the `tools` assignment with:

```ts
const tools = {
  ...existing?.tools,
  communicate: 'auto',
  approval_response: 'auto',
  list_conversations: 'auto',
  get_conversation: 'auto',
  query_usage: 'auto',
  compact_conversation: 'auto',
};
```

- [ ] **Step 4: Run the tests and typecheck**

Run: `npx vitest run src/inbound.test.ts && npm run typecheck`
Expected: PASS, no type errors.

- [ ] **Step 5: Commit**

```bash
git add src/inbound.ts src/inbound.test.ts
git commit -m "feat(teams): grant conversation tools to auto-provisioned participants"
```

---

### Task 11: Register `card.action.switch`

**Files:**
- Modify: `src/bot-pool.ts`
- Modify: `src/bot-pool.test.ts`

**Interfaces:**
- Produces: `AppHandlers.onConversationAction: (entry: BotEntry, ctx: never) => Promise<CardInvokeResponse>`
- Registers: `card.action.switch`

- [ ] **Step 1: Write the failing test**

In `src/bot-pool.test.ts`, add the handler to the `handlers` constant:

```ts
onConversationAction: async () => ({ status: 200, body: {
  statusCode: 200, type: 'application/vnd.microsoft.activity.message', value: 'Switch recorded.',
} }),
```

Add the test:

```ts
it('returns the invoke response from the SDK switch action callback', async () => {
  const on = vi.spyOn(App.prototype, 'on');
  const entry = bot('assistant', 'app');
  const onConversationAction = vi.fn(handlers.onConversationAction);
  try {
    defaultAppFactory(
      entry,
      { ...handlers, onConversationAction },
      { kind: 'http', expressApp: express() },
    );
    const callback = on.mock.calls.find(([name]) => name === 'card.action.switch')?.[1];
    expect(callback).toBeDefined();
    const context = { activity: {} } as never;
    expect(await callback!(context)).toEqual({
      status: 200,
      body: {
        statusCode: 200,
        type: 'application/vnd.microsoft.activity.message',
        value: 'Switch recorded.',
      },
    });
    expect(onConversationAction).toHaveBeenCalledWith(entry, context);
  } finally {
    on.mockRestore();
  }
});
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `npx vitest run src/bot-pool.test.ts -t "switch"`
Expected: FAIL — the callback is not registered.

- [ ] **Step 3: Implement the registration**

In `src/bot-pool.ts`, add to `AppHandlers`:

```ts
onConversationAction: (entry: BotEntry, ctx: never) => Promise<CardInvokeResponse>;
```

In `defaultAppFactory`, after the tool action registration, add:

```ts
app.on('card.action.switch', (ctx) => handlers.onConversationAction(entry, ctx as never));
```

To keep the tree typechecking before the real handler lands in Task 12, add the handler to `BotPool` in `src/connector.ts` `start`:

```ts
onConversationAction: (bot, context) => this.onConversationAction(bot, context),
```

And add a temporary stub method to `TeamsConnector`:

```ts
private async onConversationAction(
  _bot: BotEntry,
  _context: TeamsContextLike,
): Promise<CardInvokeResponse> {
  return messageResponse('Switching is not available.');
}
```

- [ ] **Step 4: Run the tests and typecheck**

Run: `npx vitest run src/bot-pool.test.ts && npm run typecheck`
Expected: PASS, no type errors.

- [ ] **Step 5: Commit**

```bash
git add src/bot-pool.ts src/bot-pool.test.ts
git commit -m "feat(teams): register card.action.switch"
```

---

### Task 12: Switch card action handler

**Files:**
- Modify: `src/connector.ts`
- Modify: `src/connector.test.ts`

**Interfaces:**
- Consumes: `resolveChatParticipant`, `listConversations`, `bindConversation`, `conversationCard`, `switchedCard`, `conversationErrorCard`
- Produces: `TeamsConnector.onConversationAction`
- Produces: private `queueKeyFor(bot, activity): string`

- [ ] **Step 1: Write the failing tests**

Add to `src/connector.test.ts` inside the `describe('TeamsConnector')` block:

```ts
it('switches the chat to a chosen conversation and updates the card', async () => {
  const { app, deps, updates } = setup();
  const factory = vi.fn<TeamsAppFactory>(() => app);
  const dir = await mkdtemp(join(tmpdir(), 'teams-switch-'));
  const connector = new TeamsConnector({ stateDir: dir, transport: 'socket' }, deps, factory);
  const callTool = vi.fn(async () => ({
    result: { status: 'success', data: { conversations: [
      { id: 'legion-1', resolvedTitle: 'First', messageCount: 4,
        updatedAt: '2026-11-24T00:00:00.000Z', status: 'active' },
      { id: 'legion-2', resolvedTitle: 'Second', messageCount: 2,
        updatedAt: '2026-11-25T00:00:00.000Z', status: 'active' },
    ] } },
    conversationId: 'c',
  }));
  await connector.start({ callTool } as never);
  try {
    connector['state']!.legionConversations['app:alice'] = 'legion-1';
    connector['state']!.legionConversationTargets['legion-1'] = 'app:alice';
    const [entry, handlers] = factory.mock.calls[0];
    const response = await handlers.onConversationAction(entry, {
      send: vi.fn(),
      activity: {
        from: { aadObjectId: 'aad-1' },
        channelData: { tenant: { id: 't' } },
        conversation: { conversationType: 'personal', id: 'conv-1' },
        replyToId: 'card-1',
        value: { action: { data: {
          action: 'switch', conversationId: 'legion-2', key: 'app:alice',
        } } },
      },
    } as never);
    expect(response).toMatchObject({
      status: 200,
      body: { statusCode: 200, type: 'application/vnd.microsoft.card.adaptive' },
    });
    expect(JSON.stringify(response.body)).toContain('Running switch');
    await vi.waitFor(() => expect(updates).toHaveLength(1));
    expect(updates[0]).toMatchObject({ conversationId: 'conv-1', activityId: 'card-1' });
    expect(JSON.stringify(updates[0].activity)).toContain('Switched conversation');
    expect(connector['state']!.legionConversations['app:alice']).toBe('legion-2');
    expect(connector['state']!.legionConversationTargets['legion-2']).toBe('app:alice');
    expect(connector['state']!.legionConversationTargets['legion-1']).toBeUndefined();
  } finally {
    await connector.stop();
  }
});

it('pages the picker without changing state', async () => {
  const { app, deps, updates } = setup();
  const factory = vi.fn<TeamsAppFactory>(() => app);
  const dir = await mkdtemp(join(tmpdir(), 'teams-switch-page-'));
  const connector = new TeamsConnector({ stateDir: dir, transport: 'socket' }, deps, factory);
  const conversations = Array.from({ length: 7 }, (_, index) => ({
    id: `legion-${index}`,
    resolvedTitle: `Conversation ${index}`,
    messageCount: 1,
    updatedAt: `2026-11-${String(20 + index).padStart(2, '0')}T00:00:00.000Z`,
    status: 'active',
  }));
  const callTool = vi.fn(async () => ({
    result: { status: 'success', data: { conversations } },
    conversationId: 'c',
  }));
  await connector.start({ callTool } as never);
  try {
    connector['state']!.legionConversations['app:alice'] = 'legion-0';
    const [entry, handlers] = factory.mock.calls[0];
    await handlers.onConversationAction(entry, {
      send: vi.fn(),
      activity: {
        from: { aadObjectId: 'aad-1' },
        channelData: { tenant: { id: 't' } },
        conversation: { conversationType: 'personal', id: 'conv-1' },
        replyToId: 'card-1',
        value: { action: { data: { action: 'switch', offset: 5, key: 'app:alice' } } },
      },
    } as never);
    await vi.waitFor(() => expect(updates).toHaveLength(1));
    expect(JSON.stringify(updates[0].activity)).toContain('Showing 6–7 of 7');
    expect(connector['state']!.legionConversations['app:alice']).toBe('legion-0');
  } finally {
    await connector.stop();
  }
});

it('shows an error card for a conversation outside the active list', async () => {
  const { app, deps, updates } = setup();
  const factory = vi.fn<TeamsAppFactory>(() => app);
  const dir = await mkdtemp(join(tmpdir(), 'teams-switch-missing-'));
  const connector = new TeamsConnector({ stateDir: dir, transport: 'socket' }, deps, factory);
  const callTool = vi.fn(async () => ({
    result: { status: 'success', data: { conversations: [] } },
    conversationId: 'c',
  }));
  await connector.start({ callTool } as never);
  try {
    const [entry, handlers] = factory.mock.calls[0];
    await handlers.onConversationAction(entry, {
      send: vi.fn(),
      activity: {
        from: { aadObjectId: 'aad-1' },
        channelData: { tenant: { id: 't' } },
        conversation: { conversationType: 'personal', id: 'conv-1' },
        replyToId: 'card-1',
        value: { action: { data: {
          action: 'switch', conversationId: 'legion-x', key: 'app:alice',
        } } },
      },
    } as never);
    await vi.waitFor(() => expect(updates).toHaveLength(1));
    expect(JSON.stringify(updates[0].activity)).toContain('Conversation not found.');
  } finally {
    await connector.stop();
  }
});

it('rejects an invalid switch payload without listing conversations', async () => {
  const { app, deps, updates } = setup();
  const factory = vi.fn<TeamsAppFactory>(() => app);
  const dir = await mkdtemp(join(tmpdir(), 'teams-switch-invalid-'));
  const connector = new TeamsConnector({ stateDir: dir, transport: 'socket' }, deps, factory);
  const callTool = vi.fn(async () => ({ result: { status: 'success' }, conversationId: 'c' }));
  await connector.start({ callTool } as never);
  try {
    const [entry, handlers] = factory.mock.calls[0];
    const invoke = (data: Record<string, unknown>) => handlers.onConversationAction(entry, {
      send: vi.fn(),
      activity: {
        from: { aadObjectId: 'aad-1' },
        channelData: { tenant: { id: 't' } },
        conversation: { conversationType: 'personal', id: 'conv-1' },
        replyToId: 'card-1',
        value: { action: { data } },
      },
    } as never);
    expect(await invoke({ action: 'switch' })).toHaveProperty('body.value', 'Invalid switch payload.');
    expect(await invoke({ action: 'switch', key: 'app:bob', conversationId: 'c1' }))
      .toHaveProperty('body.value', 'Invalid switch payload.');
    expect(await invoke({ action: 'switch', conversationId: 'x'.repeat(257) }))
      .toHaveProperty('body.value', 'Invalid switch payload.');
    expect(await invoke({ action: 'switch', offset: '5' }))
      .toHaveProperty('body.value', 'Invalid switch payload.');
    expect(callTool).not.toHaveBeenCalled();
    expect(updates).toHaveLength(0);
  } finally {
    await connector.stop();
  }
});

it('rejects tenant and conversation-type mismatches for switch', async () => {
  const { app, deps, updates } = setup();
  const factory = vi.fn<TeamsAppFactory>(() => app);
  const dir = await mkdtemp(join(tmpdir(), 'teams-switch-scope-'));
  const connector = new TeamsConnector({ stateDir: dir, transport: 'socket' }, deps, factory);
  const callTool = vi.fn(async () => ({ result: { status: 'success' }, conversationId: 'c' }));
  await connector.start({ callTool } as never);
  try {
    const [entry, handlers] = factory.mock.calls[0];
    const invoke = (tenant: string, conversationType: string) => handlers.onConversationAction(entry, {
      send: vi.fn(),
      activity: {
        from: { aadObjectId: 'aad-1' },
        channelData: { tenant: { id: tenant } },
        conversation: { conversationType, id: 'conv-1' },
        value: { action: { data: { action: 'switch', offset: 0 } } },
      },
    } as never);
    expect(await invoke('other', 'personal'))
      .toHaveProperty('body.value', 'This bot is not available in your tenant.');
    expect(await invoke('t', 'channel'))
      .toHaveProperty('body.value', 'Only personal chats and group chats are supported.');
    expect(callTool).not.toHaveBeenCalled();
    expect(updates).toHaveLength(0);
  } finally {
    await connector.stop();
  }
});

it('rejects unregistered personal and group clickers', async () => {
  const { app, deps } = setup();
  const factory = vi.fn<TeamsAppFactory>(() => app);
  const dir = await mkdtemp(join(tmpdir(), 'teams-switch-unknown-'));
  const connector = new TeamsConnector({ stateDir: dir, transport: 'socket' }, deps, factory);
  const callTool = vi.fn(async () => ({ result: { status: 'success' }, conversationId: 'c' }));
  await connector.start({ callTool } as never);
  try {
    const [entry, handlers] = factory.mock.calls[0];
    const invoke = (conversationType: string, id: string) => handlers.onConversationAction(entry, {
      send: vi.fn(),
      activity: {
        from: { aadObjectId: 'aad-9' },
        channelData: { tenant: { id: 't' } },
        conversation: { conversationType, id },
        value: { action: { data: { action: 'switch', offset: 0 } } },
      },
    } as never);
    expect(await invoke('personal', 'conv-1')).toHaveProperty('body.value', 'Not registered.');
    expect(await invoke('groupChat', '19:unknown@thread.v2'))
      .toHaveProperty('body.value', 'This group is not registered.');
    expect(callTool).not.toHaveBeenCalled();
  } finally {
    await connector.stop();
  }
});

it('switches a group chat under the group key', async () => {
  const { app, deps, updates } = setup();
  const factory = vi.fn<TeamsAppFactory>(() => app);
  const dir = await mkdtemp(join(tmpdir(), 'teams-switch-group-'));
  const connector = new TeamsConnector({ stateDir: dir, transport: 'socket' }, deps, factory);
  const callTool = vi.fn(async () => ({
    result: { status: 'success', data: { conversations: [{
      id: 'legion-group-2', resolvedTitle: 'Group thread', messageCount: 2,
      updatedAt: '2026-11-25T00:00:00.000Z', status: 'active',
    }] } },
    conversationId: 'c',
  }));
  await connector.start({ callTool } as never);
  try {
    connector['state']!.legionConversations['app:team-eng:19:g@thread.v2'] = 'legion-group-1';
    const [entry, handlers] = factory.mock.calls[0];
    await handlers.onConversationAction(entry, {
      send: vi.fn(),
      activity: {
        from: { aadObjectId: 'aad-1' },
        channelData: { tenant: { id: 't' } },
        conversation: { conversationType: 'groupChat', id: '19:g@thread.v2' },
        replyToId: 'card-1',
        value: { action: { data: {
          action: 'switch', conversationId: 'legion-group-2',
          key: 'app:team-eng:19:g@thread.v2',
        } } },
      },
    } as never);
    await vi.waitFor(() => expect(updates).toHaveLength(1));
    expect(callTool).toHaveBeenCalledWith('team-eng', 'list_conversations', expect.anything());
    expect(connector['state']!.legionConversations['app:team-eng:19:g@thread.v2'])
      .toBe('legion-group-2');
  } finally {
    await connector.stop();
  }
});

it('uses the fallback participant for an unknown personal clicker', async () => {
  const { app, deps, participants, updates } = setup();
  participants.push({
    id: 'operator', name: 'Operator', type: 'user', tools: {}, status: 'active',
    identities: [{ connector: 'teams', externalId: 'aad-op' }],
  });
  const factory = vi.fn<TeamsAppFactory>(() => app);
  const dir = await mkdtemp(join(tmpdir(), 'teams-switch-fallback-'));
  const connector = new TeamsConnector({
    stateDir: dir,
    transport: 'socket',
    unknownUserPolicy: 'fallback',
    fallbackUserId: 'operator',
  }, deps, factory);
  const callTool = vi.fn(async () => ({
    result: { status: 'success', data: { conversations: [{
      id: 'legion-op', resolvedTitle: 'Fallback thread', messageCount: 1,
      updatedAt: '2026-11-25T00:00:00.000Z', status: 'active',
    }] } },
    conversationId: 'c',
  }));
  await connector.start({ callTool } as never);
  try {
    const [entry, handlers] = factory.mock.calls[0];
    await handlers.onConversationAction(entry, {
      send: vi.fn(),
      activity: {
        from: { aadObjectId: 'aad-9' },
        channelData: { tenant: { id: 't' } },
        conversation: { conversationType: 'personal', id: 'conv-1' },
        replyToId: 'card-1',
        value: { action: { data: {
          action: 'switch', conversationId: 'legion-op', key: 'app:operator:aad-9',
        } } },
      },
    } as never);
    await vi.waitFor(() => expect(updates).toHaveLength(1));
    expect(callTool).toHaveBeenCalledWith('operator', 'list_conversations', expect.anything());
    expect(connector['state']!.legionConversations['app:operator:aad-9']).toBe('legion-op');
  } finally {
    await connector.stop();
  }
});

it('sends a follow-up card when replyToId is absent', async () => {
  const { app, deps, sent } = setup();
  const factory = vi.fn<TeamsAppFactory>(() => app);
  const dir = await mkdtemp(join(tmpdir(), 'teams-switch-followup-'));
  const connector = new TeamsConnector({ stateDir: dir, transport: 'socket' }, deps, factory);
  const callTool = vi.fn(async () => ({
    result: { status: 'success', data: { conversations: [{
      id: 'legion-2', resolvedTitle: 'Second', messageCount: 1,
      updatedAt: '2026-11-25T00:00:00.000Z', status: 'active',
    }] } },
    conversationId: 'c',
  }));
  await connector.start({ callTool } as never);
  try {
    const [entry, handlers] = factory.mock.calls[0];
    await handlers.onConversationAction(entry, {
      send: vi.fn(),
      activity: {
        from: { aadObjectId: 'aad-1' },
        channelData: { tenant: { id: 't' } },
        conversation: { conversationType: 'personal', id: 'conv-1' },
        value: { action: { data: {
          action: 'switch', conversationId: 'legion-2', key: 'app:alice',
        } } },
      },
    } as never);
    await vi.waitFor(() => expect(sent).toHaveLength(1));
    expect(sent[0].conversationId).toBe('conv-1');
    expect(JSON.stringify(sent[0].activity)).toContain('Switched conversation');
  } finally {
    await connector.stop();
  }
});

it('serializes a switch behind an in-flight turn for the same chat', async () => {
  const { app, deps, updates } = setup();
  const factory = vi.fn<TeamsAppFactory>(() => app);
  const dir = await mkdtemp(join(tmpdir(), 'teams-switch-queue-'));
  const connector = new TeamsConnector({ stateDir: dir, transport: 'socket' }, deps, factory);
  const callTool = vi.fn(async () => ({
    result: { status: 'success', data: { conversations: [{
      id: 'legion-2', resolvedTitle: 'Second', messageCount: 1,
      updatedAt: '2026-11-25T00:00:00.000Z', status: 'active',
    }] } },
    conversationId: 'c',
  }));
  await connector.start({ callTool } as never);
  try {
    let release!: () => void;
    const gate = new Promise<void>((resolve) => { release = resolve; });
    const turn = connector['queue'].run('app:aad-1', async () => { await gate; });
    const [entry, handlers] = factory.mock.calls[0];
    await handlers.onConversationAction(entry, {
      send: vi.fn(),
      activity: {
        from: { aadObjectId: 'aad-1' },
        channelData: { tenant: { id: 't' } },
        conversation: { conversationType: 'personal', id: 'conv-1' },
        replyToId: 'card-1',
        value: { action: { data: {
          action: 'switch', conversationId: 'legion-2', key: 'app:alice',
        } } },
      },
    } as never);
    expect(updates).toHaveLength(0);
    release();
    await turn;
    await vi.waitFor(() => expect(updates).toHaveLength(1));
  } finally {
    await connector.stop();
  }
});
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `npx vitest run src/connector.test.ts -t "switch"`
Expected: FAIL — `handlers.onConversationAction` is undefined.

- [ ] **Step 3: Implement the handler**

In `src/connector.ts`, add the conversations import and extend the existing identity import:

```ts
import {
  bindConversation,
  conversationCard,
  conversationErrorCard,
  listConversations,
  switchedCard,
} from './conversations.js';
import {
  findGroupActor,
  findTeamsActor,
  readTeamsId,
  resolveChatParticipant,
} from './identity.js';
```

Add constants near `MAX_TOOL_NAME`:

```ts
/** Longest conversation id or key accepted from an untrusted card payload. */
const MAX_CONVERSATION_VALUE = 256;
```

Add the queue-key helper and use it in `onMessage`:

```ts
private queueKeyFor(bot: BotEntry, activity: TeamsContextLike['activity']): string {
  return activity.conversation?.conversationType === 'groupChat'
    ? `${bot.appId}:${activity.conversation?.id || 'unknown'}`
    : `${bot.appId}:${readTeamsId(activity.from) || 'unknown'}`;
}
```

In `onMessage`, replace the key computation with:

```ts
const key = this.queueKeyFor(bot, activity);
```

In `start`, add to the `BotPool` handlers:

```ts
onConversationAction: (bot, context) => this.onConversationAction(bot, context),
```

Replace the temporary `onConversationAction` stub from Task 11 with the full handler:

```ts
private async onConversationAction(
  bot: BotEntry,
  context: TeamsContextLike,
): Promise<CardInvokeResponse> {
  if (!this.ctx || !this.state || !this.store || !this.pool) {
    return messageResponse('This bot is not available.');
  }
  if (!isBotTenant(bot, context.activity)) return messageResponse(tenantMismatchText);
  const conversationType = context.activity.conversation?.conversationType;
  if (conversationType !== 'personal' && !(conversationType === 'groupChat' && this.options.groupChats)) {
    return messageResponse(this.options.groupChats
      ? 'Only personal chats and group chats are supported.'
      : 'Only personal chats are supported.');
  }

  const isGroup = conversationType === 'groupChat';
  const teamsConversationId = context.activity.conversation?.id?.trim();
  const aadObjectId = readTeamsId(context.activity.from);
  const participant = resolveChatParticipant({
    collective: this.deps.collective,
    appId: bot.appId,
    aadObjectId,
    conversationId: teamsConversationId,
    isGroup,
    options: this.options,
  });
  if (!participant) {
    return messageResponse(isGroup ? 'This group is not registered.' : 'Not registered.');
  }

  const data = (context.activity.value?.action?.data ?? {}) as {
    action?: unknown;
    conversationId?: unknown;
    offset?: unknown;
    key?: unknown;
  };
  if (data.action !== 'switch') return messageResponse('Invalid switch payload.');

  const payloadKey = typeof data.key === 'string' ? data.key.trim() : '';
  if (payloadKey.length > MAX_CONVERSATION_VALUE) return messageResponse('Invalid switch payload.');
  if (payloadKey && payloadKey !== participant.key) return messageResponse('Invalid switch payload.');
  const stateKey = payloadKey || participant.key;

  const rawConversationId =
    typeof data.conversationId === 'string' ? data.conversationId.trim() : '';
  if (rawConversationId.length > MAX_CONVERSATION_VALUE) {
    return messageResponse('Invalid switch payload.');
  }
  const conversationId = rawConversationId || undefined;

  const rawOffset = data.offset;
  if (rawOffset !== undefined && (typeof rawOffset !== 'number' || !Number.isFinite(rawOffset))) {
    return messageResponse('Invalid switch payload.');
  }
  const offset = rawOffset === undefined ? undefined : Math.max(0, Math.floor(rawOffset));
  if (conversationId === undefined && offset === undefined) {
    return messageResponse('Invalid switch payload.');
  }
  if (!teamsConversationId) return messageResponse('Missing conversation.');

  const app = this.pool.getByAppId(bot.appId);
  if (!app) return messageResponse('Teams app is not running.');

  const cardActivityId = context.activity.replyToId;
  const ctx = this.ctx;
  const state = this.state;
  const store = this.store;
  const queueKey = this.queueKeyFor(bot, context.activity);

  void this.queue.run(queueKey, async () => {
    let card: IAdaptiveCard;
    try {
      const entries = await listConversations(ctx, participant.participantId);
      if (conversationId === undefined) {
        card = conversationCard(entries, state.legionConversations[stateKey], offset ?? 0, stateKey);
      } else {
        const entry = entries.find((candidate) => candidate.id === conversationId);
        if (!entry) {
          card = conversationErrorCard('Conversation not found.');
        } else if (entry.id === state.legionConversations[stateKey]) {
          card = switchedCard(entry, true);
        } else {
          bindConversation(state, stateKey, entry.id);
          await store.save(state);
          card = switchedCard(entry);
        }
      }
    } catch (error) {
      console.error('Teams switch action failed', describeError(error));
      card = conversationErrorCard('Could not switch conversation.');
    }

    await delay(ACK_SETTLE_MS);
    try {
      const activity = new MessageActivityInput().addCard('adaptive', card);
      if (cardActivityId) {
        try {
          await app.api.conversations.updateActivity(teamsConversationId, cardActivityId, activity);
        } catch (error) {
          console.error('Teams card update failed', describeError(error));
          await app.send(teamsConversationId, activity);
        }
      } else {
        await app.send(teamsConversationId, activity);
      }
    } catch (error) {
      console.error('Teams switch card delivery failed', describeError(error));
    }
  });

  return cardResponse(runningCard('switch'));
}
```

- [ ] **Step 4: Run the tests and typecheck**

Run: `npx vitest run src/connector.test.ts src/bot-pool.test.ts && npm run typecheck`
Expected: PASS, no type errors.

- [ ] **Step 5: Commit**

```bash
git add src/connector.ts src/connector.test.ts
git commit -m "feat(teams): handle card switches with a running ack"
```

---

### Task 13: Teams app manifest

**Files:**
- Modify: `scripts/provision-teams-bot.mjs`
- Modify: `src/provision.test.ts`

**Interfaces:**
- Produces: manifest `version: '1.1.0'` and `commandLists` for `personal` and `groupChat`.

- [ ] **Step 1: Write the failing test**

In `src/provision.test.ts`, after the existing `manifest.bots` assertion, add:

```ts
expect(manifest.version).toBe('1.1.0');
expect(manifest.commandLists).toEqual([
  {
    scopes: ['personal', 'groupChat'],
    commands: [
      { title: 'new', description: 'Start a new conversation' },
      { title: 'switch', description: 'List and switch to an existing conversation' },
      { title: 'compact', description: 'Compact this conversation' },
      { title: 'context', description: 'Show usage and context details' },
      { title: 'whoami', description: 'Show identity and chat details' },
      { title: 'pair', description: 'Pair a pending Teams identity' },
      { title: 'help', description: 'List commands' },
    ],
  },
]);
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `npx vitest run src/provision.test.ts`
Expected: FAIL — the manifest version is `1.0.0` and there is no `commandLists`.

- [ ] **Step 3: Implement the manifest change**

In `scripts/provision-teams-bot.mjs`, change `version: '1.0.0'` to `version: '1.1.0'`, and add after `permissions`:

```js
commandLists: [
  {
    scopes: ['personal', 'groupChat'],
    commands: [
      { title: 'new', description: 'Start a new conversation' },
      { title: 'switch', description: 'List and switch to an existing conversation' },
      { title: 'compact', description: 'Compact this conversation' },
      { title: 'context', description: 'Show usage and context details' },
      { title: 'whoami', description: 'Show identity and chat details' },
      { title: 'pair', description: 'Pair a pending Teams identity' },
      { title: 'help', description: 'List commands' },
    ],
  },
],
```

- [ ] **Step 4: Run the tests and typecheck**

Run: `npx vitest run src/provision.test.ts && npm run typecheck`
Expected: PASS, no type errors.

- [ ] **Step 5: Commit**

```bash
git add scripts/provision-teams-bot.mjs src/provision.test.ts
git commit -m "feat(teams): advertise conversation commands in the app manifest"
```

---

### Task 14: Documentation

**Files:**
- Modify: `README.md`
- Modify: `docs/deploy.md`

- [ ] **Step 1: Update README**

In the options table, add:

```markdown
| `compaction` | unset | `{ agentId, instruction? }` for `/compact`; `agentId` is the summarizer agent and `instruction` overrides the default summarization instruction. |
```

In the "Chat commands" section, add:

```markdown
- `/switch` — list the chat participant's active Legion conversations as an Adaptive Card; choose one to switch this chat to it. In a personal chat, a shared fallback participant may list threads from other fallback chats; choosing one still rebinds only this Teams chat.
- `/switch <conversation-id>` — switch directly to a listed active conversation.
- `/compact [instruction]` — compact the current conversation through the configured summarizer agent; the optional instruction overrides the default.
- `/context` — show the current conversation title, status, message counts, and token/cost totals.
```

Update the existing `/new` bullet to note it does not delete the previous conversation:

```markdown
- `/new` — start a new Legion conversation for the current chat (the group, in a group chat); the previous conversation remains available to `/switch`.
```

In "Identity and credentials", change the auto-provision grant sentence to:

```markdown
It grants the provisioned user `communicate`, `approval_response`, `list_conversations`, `get_conversation`, `query_usage`, and `compact_conversation`.
```

Add after the tools grants snippet:

```markdown
Existing paired users and groups are not auto-provisioned again. To enable `/switch`, `/context`, and `/compact` for an existing participant, grant the four conversation tools:

```json
{
  "tools": {
    "list_conversations": "auto",
    "get_conversation": "auto",
    "query_usage": "auto",
    "compact_conversation": "auto"
  }
}
```

Commands fail closed when a grant is missing: the core tool error appears in the reply.
```

In "Provisioning", add:

```markdown
The manifest is version `1.1.0` and includes a `commandLists` entry for personal and group chats. Existing installs do not show the command list until the app package is updated in Teams; rerun the provisioning script for each bot and upload or update the generated package.
```

- [ ] **Step 2: Update deploy docs**

In `docs/deploy.md`, after the provisioning paragraph, add:

```markdown
The generated Teams app manifest is version `1.1.0` and advertises the `/new`, `/switch`, `/compact`, `/context`, `/whoami`, `/pair`, and `/help` commands. Updating an existing install requires uploading the new package (or updating the app) in Teams; the command list is not picked up from a bot restart alone.
```

- [ ] **Step 3: Verify docs render and tests still pass**

Run: `npm test && npm run typecheck`
Expected: PASS.

- [ ] **Step 4: Commit**

```bash
git add README.md docs/deploy.md
git commit -m "docs(teams): document conversation commands and grants"
```

---

### Task 15: Full verification and live check

**Files:** none (verification only).

- [ ] **Step 1: Run the full suite**

```bash
npm test
npm run typecheck
npm run build
```

Expected: all tests pass, no type errors, build succeeds.

- [ ] **Step 2: Review the diff for accidental core changes**

```bash
git diff --stat main...HEAD
git diff main...HEAD -- src/ | grep -n "legion/packages" || true
```

Expected: only connector files, the spec, the plan, README, deploy docs, and the provisioning script changed.

- [ ] **Step 3: Live check (operator)**

In a personal chat with a granted participant:

1. Send two messages separated by `/new`; confirm two conversations exist.
2. Run `/switch`; confirm the card lists both, with the current one marked.
3. Click the older conversation; confirm the card changes to `✅ Switched conversation` and the next message lands in that conversation.
4. Run `/context`; confirm the title, total/active counts, and token/cost totals.
5. Run `/compact`; confirm the summary message and that the reply arrives without a lingering typing indicator.

Repeat steps 2–4 in a group chat, using a group participant with the four grants. Confirm an unregistered group member can click a switch button (the group participant is the actor) but cannot switch a personal chat unless they resolve to the chat participant or fallback.

- [ ] **Step 4: Final commit (if verification changed anything)**

```bash
git status --short
```

If the tree is clean, no commit is needed. Otherwise commit the verification fix with a descriptive message.
