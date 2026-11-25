# Teams Tool Buttons Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Let an agent attach tool buttons to a Teams Adaptive Card; clicking a button runs a granted Legion tool as the clicking user and updates the card with the result.

**Architecture:** `teams__send_card` gains an optional `buttons` array and turns each entry into an `Action.Execute` action with `data.action = 'tool'`. Each bot registers `card.action.tool`. The handler resolves the clicker, acks with a running card, calls `ctx.callTool` as the clicker in the background, then updates the card in place with an agent-supplied template or a generic renderer.

**Tech Stack:** TypeScript ESM/NodeNext, Node 22, `@microsoft/teams.apps` 2.2.1, `@microsoft/teams.api` 2.2.1, `@microsoft/teams.cards` 2.2.1, `adaptivecards-templating` 2.3.1, `adaptive-expressions` 4.23.3, Vitest.

**Spec:** `docs/specs/2026-11-24-card-tool-buttons-design.md`

## Global Constraints

- Node **22.x**. Teams SDK packages pinned to **2.2.1**.
- New runtime dependencies: `adaptivecards-templating@^2.3.1` and `adaptive-expressions@^4.23.3`. The templating package requires `adaptive-expressions` as an explicit peer.
- **No Legion core changes.** Use the existing `ConnectorContext.callTool(participantId, toolName, args, opts)`.
- The **clicker is the actor**. `callTool` enforces the clicker's grants, argument scopes, and tool visibility.
- Unknown clickers are rejected with `Not registered.`; no auto-provisioning on click.
- A card has at most **6 actions** total.
- Existing `teams__send_card` behavior, approval cards, state schema **v3**, and the three `teams__*` tools must remain intact when `buttons` is absent.
- Button payloads are untrusted: validate `tool`, `args`, and `resultCard` before use.
- Every task ends with its tests passing and a commit.
- Connector test command: `npm test` (Vitest). Single file: `npx vitest run src/<file>.test.ts`.
- Typecheck: `npm run typecheck`. Build: `npm run build`.
- Work in `threads/2026-11-24-connector-tools/legion-connector-teams` on branch `feat/card-tool-buttons` off `github/main`.

## File Structure

**Created**
- `src/result-card.ts` — template expansion, generic result renderer, size fallback.
- `src/result-card.test.ts` — renderer and expansion tests.

**Modified**
- `package.json`, `package-lock.json` — templating dependencies.
- `src/render.ts` — `ToolButtonSpec`, `toolButtonActions`, `runningCard`.
- `src/render.test.ts` — button action and running card tests.
- `src/tools.ts` — `buttons` schema and validation; append actions; size check after actions.
- `src/tools.test.ts` — `buttons` tests.
- `src/bot-pool.ts` — `onToolAction` handler, `card.action.tool` registration, `updateActivity` on `TeamsAppLike`.
- `src/bot-pool.test.ts` — event registration test; fake app gains `updateActivity`.
- `src/inbound.ts` — add `replyToId?: string` to `TeamsActivity`.
- `src/connector.ts` — `onToolAction` handler, response helper renames, background card update.
- `src/connector.test.ts` — handler tests; fake app gains `updateActivity`.
- `src/tool-targets.test.ts` — fake app gains `updateActivity`.
- `README.md` — `buttons` docs, template context, clicker semantics, approval limitation.

## Setup

- [ ] Create the feature branch off the merged `main`:

```bash
git fetch github
git switch -c feat/card-tool-buttons github/main
```

- [ ] Commit the design and this plan on the branch:

```bash
git add docs/specs/2026-11-24-card-tool-buttons-design.md \
        docs/plans/2026-11-24-card-tool-buttons-implementation.md
git commit -m "docs(teams): add card tool buttons design and plan"
```

---

### Task 1: Result-card module

**Files:**
- Modify: `package.json`, `package-lock.json`
- Create: `src/result-card.ts`
- Test: `src/result-card.test.ts`

**Interfaces:**
- Produces: `ResultCardContext` (type), `genericResultCard(context)`, `expandResultCard(template, context)`, `buildResultCard(template, context, maxBytes)`.

- [ ] **Step 1: Add the templating dependencies**

Edit `package.json` dependencies to:

```json
"dependencies": {
  "@microsoft/teams.apps": "2.2.1",
  "@microsoft/teams.api": "2.2.1",
  "@microsoft/teams.cards": "2.2.1",
  "adaptive-expressions": "^4.23.3",
  "adaptivecards-templating": "^2.3.1",
  "express": "^5.3.0"
}
```

Run: `npm install`

Expected: `package-lock.json` updates and both packages are installed.

- [ ] **Step 2: Write the failing tests**

Create `src/result-card.test.ts`:

```ts
import { describe, expect, it } from 'vitest';
import { buildResultCard, expandResultCard, genericResultCard } from './result-card.js';
import type { ResultCardContext } from './result-card.js';

const base: ResultCardContext = {
  status: 'success',
  tool: 'teams__post_reaction',
  data: { reaction: 'like' },
  clicker: { id: 'alice', name: 'Alice' },
};

describe('genericResultCard', () => {
  it('renders a success header and scalar object fields as facts', () => {
    const card = genericResultCard(base);
    expect(card.body?.[0]).toMatchObject({ type: 'TextBlock', text: '✅ teams__post_reaction' });
    expect(card.body?.[1]).toMatchObject({
      type: 'FactSet',
      facts: [{ title: 'reaction', value: 'like' }],
    });
  });

  it('renders an error with the message', () => {
    const card = genericResultCard({ ...base, status: 'error', data: undefined, error: 'boom' });
    const text = JSON.stringify(card);
    expect(text).toContain('❌ teams__post_reaction');
    expect(text).toContain('boom');
  });

  it('renders an array property item by item', () => {
    const card = genericResultCard({
      ...base,
      data: { conversations: [{ name: 'Chat A' }, { name: 'Chat B' }] },
    });
    const text = JSON.stringify(card);
    expect(text).toContain('conversations');
    expect(text).toContain('Chat A');
    expect(text).toContain('Chat B');
  });

  it('renders primitive arrays as bullets', () => {
    const card = genericResultCard({ ...base, data: ['a', 'b'] });
    const text = JSON.stringify(card);
    expect(text).toContain('• a');
    expect(text).toContain('• b');
  });

  it('renders nested-only data as monospace JSON', () => {
    const card = genericResultCard({ ...base, data: { nested: { a: 1 } } });
    const text = JSON.stringify(card);
    expect(text).toContain('nested');
    expect(text).toContain('Monospace');
  });

  it('marks requires_approval errors', () => {
    const card = genericResultCard({
      ...base,
      status: 'error',
      data: undefined,
      error: 'requires_approval',
    });
    expect(JSON.stringify(card)).toContain('requires approval');
  });
});

describe('expandResultCard', () => {
  it('expands bindings against the result context', () => {
    const card = expandResultCard(
      {
        type: 'AdaptiveCard',
        version: '1.4',
        body: [{ type: 'TextBlock', text: '${clicker.name}: ${data.reaction} (${status})' }],
      },
      base,
    );
    expect(JSON.stringify(card)).toContain('Alice: like (success)');
  });

  it('expands $data arrays', () => {
    const card = expandResultCard(
      {
        type: 'AdaptiveCard',
        version: '1.4',
        body: [
          {
            type: 'Container',
            $data: '${data.conversations}',
            items: [{ type: 'TextBlock', text: '${name}' }],
          },
        ],
      },
      { ...base, data: { conversations: [{ name: 'A' }, { name: 'B' }] } },
    );
    const text = JSON.stringify(card);
    expect(text).toContain('A');
    expect(text).toContain('B');
  });

  it('returns undefined for a non-object template', () => {
    expect(expandResultCard('nope', base)).toBeUndefined();
  });

  it('returns undefined when the expanded payload is not an Adaptive Card', () => {
    expect(expandResultCard({ body: [] }, base)).toBeUndefined();
  });
});

describe('buildResultCard', () => {
  it('prefers a valid template', () => {
    const card = buildResultCard(
      {
        type: 'AdaptiveCard',
        version: '1.4',
        body: [{ type: 'TextBlock', text: '${status}' }],
      },
      base,
      10_000,
    );
    expect(JSON.stringify(card)).toContain('success');
  });

  it('falls back to the generic card for an invalid template', () => {
    const card = buildResultCard({ body: [] }, base, 10_000);
    expect(JSON.stringify(card)).toContain('✅ teams__post_reaction');
  });

  it('falls back to a compact card when over the byte limit', () => {
    const card = buildResultCard(undefined, base, 10);
    expect(JSON.stringify(card)).toContain('Result too large to display.');
  });
});
```

- [ ] **Step 3: Run the tests to verify they fail**

Run: `npx vitest run src/result-card.test.ts`

Expected: FAIL because `./result-card.js` does not exist.

- [ ] **Step 4: Implement `src/result-card.ts`**

```ts
// src/result-card.ts
import * as ACData from 'adaptivecards-templating';
import type { IAdaptiveCard } from '@microsoft/teams.cards';
import type { ToolResult } from '@legion-collective/core';

export interface ResultCardClicker {
  id: string;
  name?: string;
}

export interface ResultCardContext {
  status: ToolResult['status'];
  data?: unknown;
  error?: string;
  message?: string;
  tool: string;
  args?: unknown;
  clicker?: ResultCardClicker;
}

const MAX_TEXT = 2000;
const MAX_ITEMS = 10;
const MAX_FACTS = 12;

function truncate(value: string, max: number): string {
  return value.length <= max ? value : `${value.slice(0, max - 1)}…`;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function header(context: ResultCardContext): string {
  if (context.status === 'success') return `✅ ${context.tool}`;
  if (context.status === 'pending_approval') return `⏳ ${context.tool} requires approval`;
  if (context.status === 'rejected') return `🚫 ${context.tool}`;
  if (context.status === 'error' && /requires_approval/.test(context.error ?? '')) {
    return `⏳ ${context.tool} requires approval`;
  }
  return `❌ ${context.tool}`;
}

function scalarFacts(value: Record<string, unknown>): Array<{ title: string; value: string }> {
  const facts: Array<{ title: string; value: string }> = [];
  for (const [key, raw] of Object.entries(value)) {
    if (facts.length >= MAX_FACTS) break;
    if (typeof raw === 'string' || typeof raw === 'number' || typeof raw === 'boolean') {
      facts.push({ title: key, value: truncate(String(raw), MAX_TEXT) });
    }
  }
  return facts;
}

function arrayBody(items: unknown[]): unknown[] {
  const body: unknown[] = [];
  for (const item of items.slice(0, MAX_ITEMS)) {
    if (isRecord(item)) {
      const facts = scalarFacts(item);
      body.push(
        facts.length
          ? { type: 'FactSet', facts }
          : {
              type: 'TextBlock',
              text: truncate(JSON.stringify(item) ?? '', MAX_TEXT),
              fontType: 'Monospace',
              wrap: true,
            },
      );
    } else {
      body.push({ type: 'TextBlock', text: `• ${truncate(String(item), MAX_TEXT)}`, wrap: true });
    }
  }
  if (items.length > MAX_ITEMS) {
    body.push({
      type: 'TextBlock',
      text: `… ${items.length - MAX_ITEMS} more`,
      isSubtle: true,
      wrap: true,
    });
  }
  return body;
}

function dataBody(data: unknown): unknown[] {
  if (data === undefined || data === null) return [];
  if (typeof data === 'string') {
    return [{ type: 'TextBlock', text: truncate(data, MAX_TEXT), wrap: true }];
  }
  if (Array.isArray(data)) return arrayBody(data);
  if (isRecord(data)) {
    const body: unknown[] = [];
    const facts = scalarFacts(data);
    if (facts.length) body.push({ type: 'FactSet', facts });
    for (const [key, value] of Object.entries(data)) {
      if (!Array.isArray(value)) continue;
      body.push({ type: 'TextBlock', text: key, weight: 'Bolder', wrap: true });
      body.push(...arrayBody(value));
    }
    if (body.length === 0) {
      body.push({
        type: 'TextBlock',
        text: truncate(JSON.stringify(data) ?? '', MAX_TEXT),
        fontType: 'Monospace',
        wrap: true,
      });
    }
    return body;
  }
  return [
    { type: 'TextBlock', text: truncate(String(data), MAX_TEXT), fontType: 'Monospace', wrap: true },
  ];
}

export function genericResultCard(context: ResultCardContext): IAdaptiveCard {
  const body: unknown[] = [
    { type: 'TextBlock', text: header(context), weight: 'Bolder', wrap: true },
  ];
  if (context.status === 'error' && /requires_approval/.test(context.error ?? '')) {
    const actor = context.clicker?.name ?? context.clicker?.id ?? 'the clicking user';
    body.push({
      type: 'TextBlock',
      text: `This tool requires approval for ${actor}.`,
      wrap: true,
    });
  } else if (context.error) {
    body.push({ type: 'TextBlock', text: truncate(context.error, MAX_TEXT), wrap: true });
  }
  if (context.message) {
    body.push({ type: 'TextBlock', text: truncate(context.message, MAX_TEXT), wrap: true });
  }
  body.push(...dataBody(context.data));
  return { type: 'AdaptiveCard', version: '1.4', body } as IAdaptiveCard;
}

export function expandResultCard(
  template: unknown,
  context: ResultCardContext,
): IAdaptiveCard | undefined {
  if (!isRecord(template)) return undefined;
  try {
    const expanded = new ACData.Template(template).expand({ $root: context });
    return isRecord(expanded) && expanded.type === 'AdaptiveCard'
      ? (expanded as IAdaptiveCard)
      : undefined;
  } catch {
    return undefined;
  }
}

export function buildResultCard(
  template: unknown,
  context: ResultCardContext,
  maxBytes: number,
): IAdaptiveCard {
  const card = expandResultCard(template, context) ?? genericResultCard(context);
  if (Buffer.byteLength(JSON.stringify(card), 'utf8') <= maxBytes) return card;
  return {
    type: 'AdaptiveCard',
    version: '1.4',
    body: [
      { type: 'TextBlock', text: header(context), weight: 'Bolder', wrap: true },
      { type: 'TextBlock', text: 'Result too large to display.', wrap: true },
    ],
  } as IAdaptiveCard;
}
```

- [ ] **Step 5: Run the tests to verify they pass**

Run: `npx vitest run src/result-card.test.ts`

Expected: PASS (all tests).

- [ ] **Step 6: Typecheck**

Run: `npm run typecheck`

Expected: clean.

- [ ] **Step 7: Commit**

```bash
git add package.json package-lock.json src/result-card.ts src/result-card.test.ts
git commit -m "feat(teams): add result-card template and generic renderer"
```

---

### Task 2: Button actions and running card

**Files:**
- Modify: `src/render.ts`
- Test: `src/render.test.ts`

**Interfaces:**
- Produces: `ToolButtonSpec` (type), `toolButtonActions(buttons: ToolButtonSpec[])`, `runningCard(tool: string)`.

- [ ] **Step 1: Write the failing tests**

Append to `src/render.test.ts`:

```ts
import { runningCard, toolButtonActions } from './render.js';

describe('toolButtonActions', () => {
  it('builds Action.Execute buttons with tool payloads', () => {
    const actions = toolButtonActions([
      { title: 'Add 👍', tool: 'teams__post_reaction', args: { reaction: 'like' } },
      {
        title: 'List',
        tool: 'teams__list_conversations',
        resultCard: { type: 'AdaptiveCard', body: [] },
      },
    ]);
    expect(actions).toEqual([
      {
        type: 'Action.Execute',
        title: 'Add 👍',
        data: { action: 'tool', tool: 'teams__post_reaction', args: { reaction: 'like' } },
      },
      {
        type: 'Action.Execute',
        title: 'List',
        data: {
          action: 'tool',
          tool: 'teams__list_conversations',
          resultCard: { type: 'AdaptiveCard', body: [] },
        },
      },
    ]);
  });
});

describe('runningCard', () => {
  it('renders a running header with no actions', () => {
    const card = runningCard('teams__post_reaction');
    expect(JSON.stringify(card)).toContain('Running teams__post_reaction');
    expect(card.actions ?? []).toHaveLength(0);
  });
});
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `npx vitest run src/render.test.ts`

Expected: FAIL because `toolButtonActions` and `runningCard` are not exported.

- [ ] **Step 3: Implement the additions in `src/render.ts`**

Append after `toolCard`:

```ts
export interface ToolButtonSpec {
  title: string;
  tool: string;
  args?: unknown;
  resultCard?: unknown;
}

export function toolButtonActions(buttons: ToolButtonSpec[]): unknown[] {
  return buttons.map((button) => ({
    type: 'Action.Execute',
    title: button.title,
    data: {
      action: 'tool',
      tool: button.tool,
      ...(button.args === undefined ? {} : { args: button.args }),
      ...(button.resultCard === undefined ? {} : { resultCard: button.resultCard }),
    },
  }));
}

export function runningCard(tool: string): IAdaptiveCard {
  return {
    type: 'AdaptiveCard',
    version: '1.4',
    body: [{ type: 'TextBlock', text: `⏳ Running ${tool}…`, weight: 'Bolder', wrap: true }],
  } as IAdaptiveCard;
}
```

- [ ] **Step 4: Run the tests to verify they pass**

Run: `npx vitest run src/render.test.ts`

Expected: PASS.

- [ ] **Step 5: Typecheck**

Run: `npm run typecheck`

Expected: clean.

- [ ] **Step 6: Commit**

```bash
git add src/render.ts src/render.test.ts
git commit -m "feat(teams): build card tool buttons and running card"
```

---

### Task 3: `teams__send_card` buttons

**Files:**
- Modify: `src/tools.ts`
- Test: `src/tools.test.ts`

**Interfaces:**
- Consumes: `ToolButtonSpec`, `toolButtonActions` from Task 2.
- Produces: `teams__send_card` accepts `buttons` and appends validated actions.

- [ ] **Step 1: Write the failing tests**

Add inside `describe('teams__send_card')` in `src/tools.test.ts`:

```ts
  it('appends tool buttons to the card actions', async () => {
    const s = setup({
      conversations: {
        'app:alice': {
          conversationId: 'conv-1',
          conversationType: 'personal',
          updatedAt: '2026-01-01T00:00:00.000Z',
        },
      },
    });
    const result = await run(sendCardTool(s.deps), s.context, {
      participantId: 'alice',
      card: { type: 'AdaptiveCard', version: '1.4', body: [] },
      buttons: [
        { title: 'Add 👍', tool: 'teams__post_reaction', args: { reaction: 'like' } },
      ],
    });
    expect(result.status).toBe('success');
    const activity = s.sent[0]?.activity as {
      attachments?: Array<{ content: { actions?: unknown[] } }>;
    };
    expect(activity.attachments?.[0]?.content?.actions).toEqual([
      {
        type: 'Action.Execute',
        title: 'Add 👍',
        data: { action: 'tool', tool: 'teams__post_reaction', args: { reaction: 'like' } },
      },
    ]);
  });

  it('rejects more than six buttons', async () => {
    const s = setup({
      conversations: {
        'app:alice': {
          conversationId: 'conv-1',
          conversationType: 'personal',
          updatedAt: '2026-01-01T00:00:00.000Z',
        },
      },
    });
    const result = await run(sendCardTool(s.deps), s.context, {
      participantId: 'alice',
      card: { type: 'AdaptiveCard', version: '1.4', body: [] },
      buttons: Array.from({ length: 7 }, (_, index) => ({ title: `B${index}`, tool: 't' })),
    });
    expect(result.status).toBe('error');
    expect(result.error).toContain('at most 6');
  });

  it('rejects a button without a title or tool', async () => {
    const s = setup({
      conversations: {
        'app:alice': {
          conversationId: 'conv-1',
          conversationType: 'personal',
          updatedAt: '2026-01-01T00:00:00.000Z',
        },
      },
    });
    const noTitle = await run(sendCardTool(s.deps), s.context, {
      participantId: 'alice',
      card: { type: 'AdaptiveCard', version: '1.4', body: [] },
      buttons: [{ tool: 'teams__post_reaction' }],
    });
    expect(noTitle.status).toBe('error');
    expect(noTitle.error).toContain('title');
    const noTool = await run(sendCardTool(s.deps), s.context, {
      participantId: 'alice',
      card: { type: 'AdaptiveCard', version: '1.4', body: [] },
      buttons: [{ title: 'Go' }],
    });
    expect(noTool.status).toBe('error');
    expect(noTool.error).toContain('tool');
  });

  it('rejects non-object args and resultCard', async () => {
    const s = setup({
      conversations: {
        'app:alice': {
          conversationId: 'conv-1',
          conversationType: 'personal',
          updatedAt: '2026-01-01T00:00:00.000Z',
        },
      },
    });
    const badArgs = await run(sendCardTool(s.deps), s.context, {
      participantId: 'alice',
      card: { type: 'AdaptiveCard', version: '1.4', body: [] },
      buttons: [{ title: 'Go', tool: 't', args: [] }],
    });
    expect(badArgs.status).toBe('error');
    expect(badArgs.error).toContain('args');
    const badCard = await run(sendCardTool(s.deps), s.context, {
      participantId: 'alice',
      card: { type: 'AdaptiveCard', version: '1.4', body: [] },
      buttons: [{ title: 'Go', tool: 't', resultCard: 'nope' }],
    });
    expect(badCard.status).toBe('error');
    expect(badCard.error).toContain('resultCard');
  });

  it('rejects a card whose existing actions plus buttons exceed six', async () => {
    const s = setup({
      conversations: {
        'app:alice': {
          conversationId: 'conv-1',
          conversationType: 'personal',
          updatedAt: '2026-01-01T00:00:00.000Z',
        },
      },
    });
    const result = await run(sendCardTool(s.deps), s.context, {
      participantId: 'alice',
      card: {
        type: 'AdaptiveCard',
        version: '1.4',
        body: [],
        actions: Array.from({ length: 5 }, () => ({
          type: 'Action.Execute',
          title: 'x',
          data: {},
        })),
      },
      buttons: [
        { title: 'A', tool: 't' },
        { title: 'B', tool: 't' },
      ],
    });
    expect(result.status).toBe('error');
    expect(result.error).toContain('at most 6 actions');
  });

  it('enforces maxCardBytes after adding buttons', async () => {
    const s = setup({
      conversations: {
        'app:alice': {
          conversationId: 'conv-1',
          conversationType: 'personal',
          updatedAt: '2026-01-01T00:00:00.000Z',
        },
      },
      options: { maxCardBytes: 200 },
    });
    const result = await run(sendCardTool(s.deps), s.context, {
      participantId: 'alice',
      card: {
        type: 'AdaptiveCard',
        version: '1.4',
        body: [{ type: 'TextBlock', text: 'x'.repeat(100) }],
      },
      buttons: [{ title: 'Go', tool: 'teams__post_reaction', args: { note: 'y'.repeat(150) } }],
    });
    expect(result.status).toBe('error');
    expect(result.error).toContain('max 200');
  });
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `npx vitest run src/tools.test.ts`

Expected: FAIL because `buttons` is ignored and the size check runs before actions are added.

- [ ] **Step 3: Implement the `buttons` schema and validation in `src/tools.ts`**

Add imports:

```ts
import { toolButtonActions, type ToolButtonSpec } from './render.js';
```

Add constants and a validator above `sendCardTool`:

```ts
const MAX_CARD_ACTIONS = 6;
const MAX_BUTTON_TITLE = 40;

function validateButtons(value: unknown):
  | { ok: true; buttons: ToolButtonSpec[] }
  | { ok: false; error: string } {
  if (value === undefined) return { ok: true, buttons: [] };
  if (!Array.isArray(value)) return { ok: false, error: 'buttons must be an array' };
  if (value.length === 0) return { ok: true, buttons: [] };
  if (value.length > MAX_CARD_ACTIONS) {
    return { ok: false, error: `buttons must contain at most ${MAX_CARD_ACTIONS} entries` };
  }
  const buttons: ToolButtonSpec[] = [];
  for (const [index, raw] of value.entries()) {
    if (typeof raw !== 'object' || raw === null || Array.isArray(raw)) {
      return { ok: false, error: `buttons[${index}] must be an object` };
    }
    const entry = raw as Record<string, unknown>;
    const title = typeof entry.title === 'string' ? entry.title.trim() : '';
    if (!title) return { ok: false, error: `buttons[${index}].title is required` };
    if (title.length > MAX_BUTTON_TITLE) {
      return {
        ok: false,
        error: `buttons[${index}].title must be at most ${MAX_BUTTON_TITLE} characters`,
      };
    }
    const tool = typeof entry.tool === 'string' ? entry.tool.trim() : '';
    if (!tool) return { ok: false, error: `buttons[${index}].tool is required` };
    if (
      entry.args !== undefined &&
      (typeof entry.args !== 'object' || entry.args === null || Array.isArray(entry.args))
    ) {
      return { ok: false, error: `buttons[${index}].args must be an object` };
    }
    if (
      entry.resultCard !== undefined &&
      (typeof entry.resultCard !== 'object' ||
        entry.resultCard === null ||
        Array.isArray(entry.resultCard))
    ) {
      return { ok: false, error: `buttons[${index}].resultCard must be an object` };
    }
    buttons.push({
      title,
      tool,
      ...(entry.args === undefined ? {} : { args: entry.args }),
      ...(entry.resultCard === undefined ? {} : { resultCard: entry.resultCard }),
    });
  }
  return { ok: true, buttons };
}
```

Add `buttons` to the `parameters.properties` schema and mention buttons in the tool description:

```ts
        buttons: {
          type: 'array',
          description:
            'Optional tool buttons. Each runs a granted tool as the clicking user.',
          items: {
            type: 'object',
            properties: {
              title: { type: 'string' },
              tool: { type: 'string' },
              args: { type: 'object' },
              resultCard: { type: 'object' },
            },
            required: ['title', 'tool'],
          },
        },
```

Also change the tool description to:

```ts
    description:
      "Send an Adaptive Card to a Teams conversation through the caller's Teams bot. " +
      'Optional buttons run granted tools as the clicking user.',
```

Replace the size check and card send in `execute`:

```ts
      const checked = validateButtons(raw.buttons);
      if (!checked.ok) return fail(`teams__send_card: ${checked.error}`);

      const existingActions = Array.isArray((card as { actions?: unknown[] }).actions)
        ? (card as { actions: unknown[] }).actions
        : [];
      const actions = [...existingActions, ...toolButtonActions(checked.buttons)];
      if (checked.buttons.length > 0 && actions.length > MAX_CARD_ACTIONS) {
        return fail(
          `teams__send_card: card has ${actions.length} actions (max ${MAX_CARD_ACTIONS})`,
        );
      }
      const outgoingCard = actions.length
        ? { ...(card as Record<string, unknown>), actions }
        : card;

      const serialized = JSON.stringify(outgoingCard);
      const bytes = Buffer.byteLength(serialized, 'utf8');
      if (bytes > deps.options.maxCardBytes) {
        return fail(`teams__send_card: card is ${bytes} bytes (max ${deps.options.maxCardBytes})`);
      }

      const activity =
        typeof raw.fallbackText === 'string' && raw.fallbackText.trim()
          ? new MessageActivityInput('', { summary: raw.fallbackText })
          : new MessageActivityInput();
      activity.addCard('adaptive', outgoingCard as IAdaptiveCard);
```

- [ ] **Step 4: Run the tests to verify they pass**

Run: `npx vitest run src/tools.test.ts`

Expected: PASS.

- [ ] **Step 5: Typecheck**

Run: `npm run typecheck`

Expected: clean.

- [ ] **Step 6: Commit**

```bash
git add src/tools.ts src/tools.test.ts
git commit -m "feat(teams): add tool buttons to teams__send_card"
```

---

### Task 4: Bot-pool event and update surface

**Files:**
- Modify: `src/bot-pool.ts`, `src/bot-pool.test.ts`, `src/tools.test.ts`, `src/tool-targets.test.ts`
- Test: `src/bot-pool.test.ts`

**Interfaces:**
- Consumes: nothing from earlier tasks.
- Produces: `AppHandlers.onToolAction`, `TeamsAppLike.api.conversations.updateActivity`, `card.action.tool` registration.

- [ ] **Step 1: Write the failing test**

In `src/bot-pool.test.ts`, add `onToolAction` to the shared `handlers`:

```ts
const handlers: AppHandlers = {
  onMessage: vi.fn(async () => undefined),
  onInstall: vi.fn(async () => undefined),
  onConversationUpdate: vi.fn(async () => undefined),
  onApproval: async () => ({ status: 200, body: {
    statusCode: 200, type: 'application/vnd.microsoft.activity.message', value: 'Approval recorded.',
  } }),
  onToolAction: async () => ({ status: 200, body: {
    statusCode: 200, type: 'application/vnd.microsoft.activity.message', value: 'Tool recorded.',
  } }),
};
```

Add `updateActivity` to `fakeApp`'s `api.conversations`:

```ts
    api: {
      conversations: {
        create: async () => ({ id: 'conv' }),
        updateActivity: async () => undefined,
      },
      reactions: { add: async () => undefined, delete: async () => undefined },
    },
```

Add a test next to the approval event test:

```ts
  it('returns the invoke response from the SDK tool action callback', async () => {
    const on = vi.spyOn(App.prototype, 'on');
    const entry = bot('assistant', 'app');
    const onToolAction = vi.fn(handlers.onToolAction);
    try {
      defaultAppFactory(
        entry,
        { ...handlers, onToolAction },
        { kind: 'http', expressApp: express() },
      );
      const callback = on.mock.calls.find(([name]) => name === 'card.action.tool')?.[1];
      expect(callback).toBeDefined();
      const context = { activity: {} } as never;
      expect(await callback!(context)).toEqual({
        status: 200,
        body: {
          statusCode: 200,
          type: 'application/vnd.microsoft.activity.message',
          value: 'Tool recorded.',
        },
      });
      expect(onToolAction).toHaveBeenCalledWith(entry, context);
    } finally {
      on.mockRestore();
    }
  });
```

Also update the `TeamsAppLike` literals in `src/tools.test.ts` and `src/tool-targets.test.ts` so they compile:

```ts
    api: {
      conversations: {
        create,
        updateActivity: async () => undefined,
      },
      reactions,
    },
```

For `src/tool-targets.test.ts`:

```ts
    api: {
      conversations: {
        create,
        updateActivity: async () => undefined,
      },
      reactions: { add: async () => undefined, delete: async () => undefined },
    },
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `npx vitest run src/bot-pool.test.ts`

Expected: FAIL because `onToolAction` is not part of `AppHandlers` and `card.action.tool` is not registered.

- [ ] **Step 3: Implement the wiring in `src/bot-pool.ts`**

Add to `TeamsAppLike`:

```ts
  api: {
    conversations: {
      create(params: Record<string, unknown>): Promise<{ id: string }>;
      updateActivity(conversationId: string, id: string, params: unknown): Promise<unknown>;
    };
    reactions: {
      add(conversationId: string, activityId: string, type: string): Promise<void>;
      delete(conversationId: string, activityId: string, type: string): Promise<void>;
    };
  };
```

Add to `AppHandlers`:

```ts
  onToolAction: (entry: BotEntry, ctx: never) => Promise<CardInvokeResponse>;
```

Register the event in `defaultAppFactory` after the approval events:

```ts
  app.on('card.action.tool', (ctx) => handlers.onToolAction(entry, ctx as never));
```

- [ ] **Step 4: Run the tests to verify they pass**

Run: `npx vitest run src/bot-pool.test.ts src/tools.test.ts src/tool-targets.test.ts`

Expected: PASS.

- [ ] **Step 5: Typecheck**

Run: `npm run typecheck`

Expected: clean.

- [ ] **Step 6: Commit**

```bash
git add src/bot-pool.ts src/bot-pool.test.ts src/tools.test.ts src/tool-targets.test.ts
git commit -m "feat(teams): route card.action.tool and expose updateActivity"
```

---

### Task 5: Connector tool-action handler

**Files:**
- Modify: `src/connector.ts`, `src/inbound.ts`, `src/connector.test.ts`
- Test: `src/connector.test.ts`

**Interfaces:**
- Consumes: `runningCard` (Task 2), `buildResultCard` (Task 1), `onToolAction` (Task 4), `updateActivity` (Task 4).
- Produces: `TeamsConnector` handles `card.action.tool`.

- [ ] **Step 1: Add `replyToId` to the activity type**

In `src/inbound.ts`, add to `TeamsActivity`:

```ts
  replyToId?: string;
```

- [ ] **Step 2: Write the failing tests**

In `src/connector.test.ts`, add `updateActivity` to the `setup` app and record calls:

```ts
function setup() {
  const sent: Array<{ conversationId: string; activity: unknown }> = [];
  const updates: Array<{ conversationId: string; activityId: string; activity: unknown }> = [];
  const app: TeamsAppLike = {
    async initialize() {},
    async start() {},
    async stop() {},
    async send(conversationId, activity) {
      sent.push({ conversationId, activity });
      return { id: 'activity-1' };
    },
    api: {
      conversations: {
        create: async () => ({ id: 'conv-new' }),
        updateActivity: async (conversationId, activityId, activity) => {
          updates.push({ conversationId, activityId, activity });
        },
      },
      reactions: { add: async () => undefined, delete: async () => undefined },
    },
    on() {},
  };
  // ...participants, collective, deps unchanged...
  return { app, sent, updates, deps, participants };
}
```

Add a `describe('card tool actions')` block:

```ts
describe('card tool actions', () => {
  it('runs the tool as the clicker and updates the card in place', async () => {
    const { app, deps, updates } = setup();
    const factory = vi.fn<TeamsAppFactory>(() => app);
    const dir = await mkdtemp(join(tmpdir(), 'teams-tool-action-'));
    const connector = new TeamsConnector({ stateDir: dir, transport: 'socket' }, deps, factory);
    const callTool = vi.fn(async () => ({
      result: { status: 'success', data: { reaction: 'like' } },
      conversationId: 'c1',
    }));
    await connector.start({ callTool } as never);
    try {
      const [entry, handlers] = factory.mock.calls[0];
      const response = await handlers.onToolAction(entry, {
        send: vi.fn(),
        activity: {
          from: { aadObjectId: 'aad-1', name: 'Alice' },
          channelData: { tenant: { id: 't' } },
          conversation: { conversationType: 'personal', id: 'conv-1' },
          replyToId: 'card-1',
          value: {
            action: {
              data: {
                action: 'tool',
                tool: 'teams__post_reaction',
                args: { reaction: 'like' },
              },
            },
          },
        },
      } as never);
      expect(response).toMatchObject({
        status: 200,
        body: { statusCode: 200, type: 'application/vnd.microsoft.card.adaptive' },
      });
      expect(JSON.stringify(response.body)).toContain('Running teams__post_reaction');
      expect(callTool).toHaveBeenCalledWith(
        'alice',
        'teams__post_reaction',
        { reaction: 'like' },
        { conversationId: '' },
      );
      await vi.waitFor(() => expect(updates).toHaveLength(1));
      expect(updates[0]).toMatchObject({ conversationId: 'conv-1', activityId: 'card-1' });
      expect(JSON.stringify(updates[0].activity)).toContain('like');
    } finally {
      await connector.stop();
    }
  });

  it('rejects an unknown clicker without calling the tool', async () => {
    const { app, deps } = setup();
    const factory = vi.fn<TeamsAppFactory>(() => app);
    const dir = await mkdtemp(join(tmpdir(), 'teams-tool-unknown-'));
    const connector = new TeamsConnector({ stateDir: dir, transport: 'socket' }, deps, factory);
    const callTool = vi.fn(async () => ({ result: { status: 'success' }, conversationId: 'c' }));
    await connector.start({ callTool } as never);
    try {
      const [entry, handlers] = factory.mock.calls[0];
      const response = await handlers.onToolAction(entry, {
        send: vi.fn(),
        activity: {
          from: { aadObjectId: 'aad-2' },
          channelData: { tenant: { id: 't' } },
          conversation: { conversationType: 'personal', id: 'conv-1' },
          value: { action: { data: { action: 'tool', tool: 'teams__post_reaction' } } },
        },
      } as never);
      expect(response).toHaveProperty('body.value', 'Not registered.');
      expect(callTool).not.toHaveBeenCalled();
    } finally {
      await connector.stop();
    }
  });

  it('rejects a missing tool and malformed args', async () => {
    const { app, deps } = setup();
    const factory = vi.fn<TeamsAppFactory>(() => app);
    const dir = await mkdtemp(join(tmpdir(), 'teams-tool-invalid-'));
    const connector = new TeamsConnector({ stateDir: dir, transport: 'socket' }, deps, factory);
    const callTool = vi.fn(async () => ({ result: { status: 'success' }, conversationId: 'c' }));
    await connector.start({ callTool } as never);
    try {
      const [entry, handlers] = factory.mock.calls[0];
      const invoke = (data: unknown) => handlers.onToolAction(entry, {
        send: vi.fn(),
        activity: {
          from: { aadObjectId: 'aad-1' },
          channelData: { tenant: { id: 't' } },
          conversation: { conversationType: 'personal', id: 'conv-1' },
          value: { action: { data } },
        },
      } as never);
      expect(await invoke({ action: 'tool' })).toHaveProperty('body.value', 'Missing tool.');
      expect(await invoke({ action: 'tool', tool: 't', args: [] }))
        .toHaveProperty('body.value', 'Invalid args.');
      expect(callTool).not.toHaveBeenCalled();
    } finally {
      await connector.stop();
    }
  });

  it('rejects tenant and conversation-type mismatches', async () => {
    const { app, deps } = setup();
    const factory = vi.fn<TeamsAppFactory>(() => app);
    const dir = await mkdtemp(join(tmpdir(), 'teams-tool-scope-'));
    const connector = new TeamsConnector({ stateDir: dir, transport: 'socket' }, deps, factory);
    const callTool = vi.fn(async () => ({ result: { status: 'success' }, conversationId: 'c' }));
    await connector.start({ callTool } as never);
    try {
      const [entry, handlers] = factory.mock.calls[0];
      const invoke = (tenant: string, conversationType: string) => handlers.onToolAction(entry, {
        send: vi.fn(),
        activity: {
          from: { aadObjectId: 'aad-1' },
          channelData: { tenant: { id: tenant } },
          conversation: { conversationType, id: 'conv-1' },
          value: { action: { data: { action: 'tool', tool: 'teams__post_reaction' } } },
        },
      } as never);
      expect(await invoke('other', 'personal'))
        .toHaveProperty('body.value', 'This bot is not available in your tenant.');
      expect(await invoke('t', 'channel'))
        .toHaveProperty('body.value', 'Only personal chats and group chats are supported.');
      expect(callTool).not.toHaveBeenCalled();
    } finally {
      await connector.stop();
    }
  });

  it('renders requires_approval as a result card', async () => {
    const { app, deps, updates } = setup();
    const factory = vi.fn<TeamsAppFactory>(() => app);
    const dir = await mkdtemp(join(tmpdir(), 'teams-tool-approval-'));
    const connector = new TeamsConnector({ stateDir: dir, transport: 'socket' }, deps, factory);
    const callTool = vi.fn(async () => ({
      result: { status: 'error', error: 'requires_approval' },
      conversationId: 'c1',
    }));
    await connector.start({ callTool } as never);
    try {
      const [entry, handlers] = factory.mock.calls[0];
      await handlers.onToolAction(entry, {
        send: vi.fn(),
        activity: {
          from: { aadObjectId: 'aad-1', name: 'Alice' },
          channelData: { tenant: { id: 't' } },
          conversation: { conversationType: 'personal', id: 'conv-1' },
          replyToId: 'card-1',
          value: { action: { data: { action: 'tool', tool: 'teams__post_reaction' } } },
        },
      } as never);
      await vi.waitFor(() => expect(updates).toHaveLength(1));
      expect(JSON.stringify(updates[0].activity)).toContain('requires approval');
    } finally {
      await connector.stop();
    }
  });

  it('sends a follow-up when replyToId is absent', async () => {
    const { app, deps, sent } = setup();
    const factory = vi.fn<TeamsAppFactory>(() => app);
    const dir = await mkdtemp(join(tmpdir(), 'teams-tool-followup-'));
    const connector = new TeamsConnector({ stateDir: dir, transport: 'socket' }, deps, factory);
    const callTool = vi.fn(async () => ({
      result: { status: 'success', data: { reaction: 'like' } },
      conversationId: 'c1',
    }));
    await connector.start({ callTool } as never);
    try {
      const [entry, handlers] = factory.mock.calls[0];
      await handlers.onToolAction(entry, {
        send: vi.fn(),
        activity: {
          from: { aadObjectId: 'aad-1' },
          channelData: { tenant: { id: 't' } },
          conversation: { conversationType: 'personal', id: 'conv-1' },
          value: {
            action: {
              data: { action: 'tool', tool: 'teams__post_reaction', args: { reaction: 'like' } },
            },
          },
        },
      } as never);
      await vi.waitFor(() => expect(sent).toHaveLength(1));
      expect(sent[0].conversationId).toBe('conv-1');
      expect(JSON.stringify(sent[0].activity)).toContain('like');
    } finally {
      await connector.stop();
    }
  });

  it('reuses the mapped Legion conversation', async () => {
    const { app, deps, updates } = setup();
    const factory = vi.fn<TeamsAppFactory>(() => app);
    const dir = await mkdtemp(join(tmpdir(), 'teams-tool-mapped-'));
    const connector = new TeamsConnector({ stateDir: dir, transport: 'socket' }, deps, factory);
    const callTool = vi.fn(async () => ({
      result: { status: 'success' },
      conversationId: 'legion-1',
    }));
    await connector.start({ callTool } as never);
    try {
      connector['state']!.legionConversations['app:alice'] = 'legion-1';
      const [entry, handlers] = factory.mock.calls[0];
      await handlers.onToolAction(entry, {
        send: vi.fn(),
        activity: {
          from: { aadObjectId: 'aad-1' },
          channelData: { tenant: { id: 't' } },
          conversation: { conversationType: 'personal', id: 'conv-1' },
          replyToId: 'card-1',
          value: { action: { data: { action: 'tool', tool: 'teams__post_reaction' } } },
        },
      } as never);
      expect(callTool).toHaveBeenCalledWith(
        'alice',
        'teams__post_reaction',
        {},
        { conversationId: 'legion-1' },
      );
      await vi.waitFor(() => expect(updates).toHaveLength(1));
    } finally {
      await connector.stop();
    }
  });

  it('renders a tool exception as an error card', async () => {
    const { app, deps, updates } = setup();
    const factory = vi.fn<TeamsAppFactory>(() => app);
    const dir = await mkdtemp(join(tmpdir(), 'teams-tool-error-'));
    const connector = new TeamsConnector({ stateDir: dir, transport: 'socket' }, deps, factory);
    const callTool = vi.fn(async () => {
      throw new Error('boom');
    });
    await connector.start({ callTool } as never);
    try {
      const [entry, handlers] = factory.mock.calls[0];
      await handlers.onToolAction(entry, {
        send: vi.fn(),
        activity: {
          from: { aadObjectId: 'aad-1' },
          channelData: { tenant: { id: 't' } },
          conversation: { conversationType: 'personal', id: 'conv-1' },
          replyToId: 'card-1',
          value: { action: { data: { action: 'tool', tool: 'teams__post_reaction' } } },
        },
      } as never);
      await vi.waitFor(() => expect(updates).toHaveLength(1));
      expect(JSON.stringify(updates[0].activity)).toContain('boom');
    } finally {
      await connector.stop();
    }
  });
});
```

- [ ] **Step 3: Run the tests to verify they fail**

Run: `npx vitest run src/connector.test.ts`

Expected: FAIL because `onToolAction` is not implemented.

- [ ] **Step 4: Implement the handler in `src/connector.ts`**

Update imports:

```ts
import type {
  AnyTool,
  Connector,
  ConnectorContext,
  ConnectorRuntimeDeps,
  DeliveryOutcome,
  ToolResult,
} from '@legion-collective/core';
import { MessageActivityInput } from '@microsoft/teams.api';
import { buildResultCard } from './result-card.js';
import { decidedApprovalCard, runningCard, type ApprovalRequest } from './render.js';
```

Rename the response helpers:

```ts
function messageResponse(value: string): CardInvokeResponse {
  return { status: 200, body: { statusCode: 200, type: 'application/vnd.microsoft.activity.message', value } };
}

function cardResponse(card: ReturnType<typeof decidedApprovalCard>): CardInvokeResponse {
  return {
    status: 200,
    body: { statusCode: 200, type: 'application/vnd.microsoft.card.adaptive', value: card },
  };
}
```

Replace `approvalResponse(` with `messageResponse(` and `approvalCardResponse(` with `cardResponse(` throughout `src/connector.ts`.

Wire the handler in `start`:

```ts
      onApproval: (bot, context, decision) => this.onApproval(bot, context, decision),
      onToolAction: (bot, context) => this.onToolAction(bot, context),
```

Add the handler before `deliver`:

```ts
  private async onToolAction(
    bot: BotEntry,
    context: TeamsContextLike,
  ): Promise<CardInvokeResponse> {
    if (!this.ctx || !this.state || !this.pool) {
      return messageResponse('This bot is not available.');
    }
    if (!isBotTenant(bot, context.activity)) return messageResponse(tenantMismatchText);
    const conversationType = context.activity.conversation?.conversationType;
    if (conversationType !== 'personal' && !(conversationType === 'groupChat' && this.options.groupChats)) {
      return messageResponse(this.options.groupChats
        ? 'Only personal chats and group chats are supported.'
        : 'Only personal chats are supported.');
    }
    const aadObjectId = readTeamsId(context.activity.from);
    const participant = aadObjectId ? findTeamsActor(this.deps.collective, aadObjectId) : undefined;
    if (!participant) return messageResponse('Not registered.');

    const data = (context.activity.value?.action?.data ?? {}) as {
      tool?: unknown;
      args?: unknown;
      resultCard?: unknown;
    };
    const tool = typeof data.tool === 'string' ? data.tool.trim() : '';
    if (!tool) return messageResponse('Missing tool.');
    if (
      data.args !== undefined &&
      (typeof data.args !== 'object' || data.args === null || Array.isArray(data.args))
    ) {
      return messageResponse('Invalid args.');
    }

    const teamsConversationId = context.activity.conversation?.id;
    if (!teamsConversationId) return messageResponse('Missing conversation.');
    const cardActivityId = context.activity.replyToId;
    const app = this.pool.getByAppId(bot.appId);
    if (!app) return messageResponse('Teams app is not running.');

    const key = conversationType === 'groupChat'
      ? `${bot.appId}:${teamsConversationId}`
      : `${bot.appId}:${aadObjectId ?? 'unknown'}`;
    const legionConversationId = this.state.legionConversations[key];
    const clicker = { id: participant.id, name: participant.name };
    const args = data.args;
    const resultTemplate = data.resultCard;
    const ctx = this.ctx;
    const maxCardBytes = this.options.maxCardBytes;

    void (async () => {
      let result: ToolResult;
      try {
        const call = await ctx.callTool(
          participant.id,
          tool,
          args ?? {},
          { conversationId: legionConversationId ?? '' },
        );
        result = call.result;
      } catch (error) {
        result = { status: 'error', error: describeError(error) };
      }

      const card = buildResultCard(
        resultTemplate,
        {
          status: result.status,
          data: result.data,
          error: result.error,
          message: result.message,
          tool,
          args,
          clicker,
        },
        maxCardBytes,
      );

      try {
        const activity = new MessageActivityInput().addCard('adaptive', card);
        if (cardActivityId) {
          await app.api.conversations.updateActivity(teamsConversationId, cardActivityId, activity);
        } else {
          await app.send(teamsConversationId, activity);
        }
      } catch (error) {
        console.error('Teams card action result delivery failed', describeError(error));
      }
    })();

    return cardResponse(runningCard(tool));
  }
```

- [ ] **Step 5: Run the tests to verify they pass**

Run: `npx vitest run src/connector.test.ts`

Expected: PASS.

- [ ] **Step 6: Run the full suite and typecheck**

Run: `npm test`
Expected: PASS (existing tests plus the new ones).

Run: `npm run typecheck`
Expected: clean.

- [ ] **Step 7: Commit**

```bash
git add src/connector.ts src/inbound.ts src/connector.test.ts
git commit -m "feat(teams): run tools from card buttons as the clicker"
```

---

### Task 6: Documentation, build, and live verification

**Files:**
- Modify: `README.md`

- [ ] **Step 1: Update the `teams__send_card` row and add a tool-buttons section**

In the Tools table, change the `teams__send_card` row to mention `buttons`:

```markdown
| `teams__send_card` | `card` (required Adaptive Card object, schema 1.4+), exactly one of `participantId` or `conversationId`, optional `fallbackText`, optional `aiLabel`, optional `buttons` | Send an Adaptive Card. Cards larger than `maxCardBytes` are rejected. Buttons run granted tools as the clicking user. Returns the conversation and activity IDs. |
```

Add after the “Acting bot” paragraph:

````markdown
### Tool buttons

`teams__send_card` accepts an optional `buttons` array (max 6 entries):

```json
{
  "participantId": "teams-chris-...",
  "card": { "type": "AdaptiveCard", "version": "1.4", "body": [] },
  "buttons": [
    {
      "title": "Add 👍",
      "tool": "teams__post_reaction",
      "args": { "messageId": "m-1", "reaction": "like" }
    }
  ]
}
```

Each button becomes an `Action.Execute` with `data.action = "tool"`. Clicking it:

1. Resolves the Teams user to a Legion participant. Unknown or retired users get
   `Not registered.`; no auto-provisioning happens on click.
2. Runs `tool` as that participant through `ConnectorContext.callTool`. The clicker's
   `tools` grants, argument scopes, and tool visibility are enforced. There are no
   automatic grants.
3. Acknowledges with a “Running <tool>…” card, then updates the card in place with the
   result. If Teams omits the card message id, the result is sent as a follow-up message.

`args` is fixed at send time. The card payload is untrusted and validated before use.

**Result formatting.** A button may include a `resultCard` Adaptive Card template:

```json
{
  "title": "List chats",
  "tool": "teams__list_conversations",
  "resultCard": {
    "type": "AdaptiveCard",
    "version": "1.4",
    "body": [
      { "type": "TextBlock", "text": "Status: ${status}" },
      { "type": "TextBlock", "text": "Count: ${count(data.conversations)}" }
    ]
  }
}
```

The template is expanded with `adaptivecards-templating` against
`{status, data, error, message, tool, args, clicker}`, where `clicker` is
`{id, name?}`. `${...}` and `$data` work. If the template is missing or invalid, the
connector uses a generic renderer (facts, arrays, or monospace JSON) and falls back to a
compact card when the result exceeds `maxCardBytes`.

**Approval-gated tools.** V1 does not create approval requests from a button click. If the
clicker's policy is `requires_approval`, the card says the tool requires approval. The
approval flow for button tools is a follow-up.
````

- [ ] **Step 2: Build and run the full verification**

Run:

```bash
npm run typecheck
npm test
npm run build
```

Expected: all clean; tests pass; `dist` rebuilds.

- [ ] **Step 3: Commit the docs**

```bash
git add README.md
git commit -m "docs(teams): document card tool buttons"
```

- [ ] **Step 4: Live verification**

Rebuild the connector `dist` (Step 2 already did). Restart the running Teams-enabled Legion
instance so it loads the new connector module, then test:

1. In the live workspace's participant file for the clicking user (for example
   `.legion/collective/participants/teams-chris-willingham-1j2gcm.json`), grant the button
   tool with `"teams__post_reaction": "auto"`. The clicker's grants are the authority; the
   sender's grants are irrelevant to the click.
2. Ask the Assistant in Teams to send a card with a `teams__post_reaction` button.
3. Click the button as a registered user.
4. Verify the card changes to “Running…” and then to a result card, and that the reaction
   lands on the target message.

Record the observed result in the PR description.

- [ ] **Step 5: Push and open the PR**

```bash
git push -u github feat/card-tool-buttons
gh pr create --repo Prograde-Labs/legion-connector-teams --base main \
  --head feat/card-tool-buttons \
  --title "feat: tool buttons on Adaptive Cards" \
  --body "Adds optional \`buttons\` to \`teams__send_card\`. Clicking runs a granted tool as the clicking user and updates the card with a templated or generic result. Design: docs/specs/2026-11-24-card-tool-buttons-design.md. Plan: docs/plans/2026-11-24-card-tool-buttons-implementation.md."
```

---

## Self-review notes

- **Spec coverage:** `buttons` API, clicker actor, ack/update, template + generic
  fallback, fail-closed behavior, max 6 actions, `maxCardBytes`, no core changes, state
  unchanged, docs, and live test are each covered by a task.
- **Type consistency:** `ResultCardContext`, `ToolButtonSpec`, `toolButtonActions`,
  `runningCard`, `buildResultCard`, `onToolAction`, and `updateActivity` are introduced
  once and reused with the same names.
- **Placeholders:** none; every code and test step contains concrete code.
