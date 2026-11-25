import { describe, expect, it } from 'vitest';
import {
  approvalCard,
  chunkMarkdown,
  decidedApprovalCard,
  formatToolResult,
  runningCard,
  toolButtonActions,
  toolCard,
} from './render.js';

describe('render', () => {
  it('returns a single chunk for short text', () => {
    expect(chunkMarkdown('hello', 100)).toEqual(['hello']);
  });

  it('clamps nonpositive chunk sizes and preserves content', () => {
    expect(chunkMarkdown('abc', 0).join('')).toBe('abc');
    expect(chunkMarkdown('abc', -2).join('')).toBe('abc');
  });

  it('normalizes NaN to a chunk size of one', () => {
    expect(chunkMarkdown('abc', NaN)).toEqual(['a', 'b', 'c']);
  });

  it('splits long text on a safe boundary and preserves content', () => {
    const text = `${'a'.repeat(40)}\n${'b'.repeat(40)}\n${'c'.repeat(40)}`;
    const chunks = chunkMarkdown(text, 50);
    expect(chunks.length).toBeGreaterThan(1);
    expect(chunks.join('')).toBe(text);
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
    expect(card.body).toContainEqual(expect.objectContaining({
      type: 'Input.Text',
      id: 'reason',
    }));
  });

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

  it('builds a decided approval card without actions', () => {
    const card = decidedApprovalCard({
      outcome: 'approved',
      tool: 'file_write',
      argsText: '{"path":"test.txt"}',
      requesterId: 'assistant',
    });
    expect(card.type).toBe('AdaptiveCard');
    expect(card.actions ?? []).toHaveLength(0);
    expect(JSON.stringify(card)).toContain('Approved');
    expect(JSON.stringify(card)).toContain('file_write');
  });

  it('builds a tool card', () => {
    const card = toolCard('communicate', '{"to":"alice"}');
    expect(card.type).toBe('AdaptiveCard');
    expect(JSON.stringify(card)).toContain('communicate');
  });

  it('uses a fallback for circular approval arguments', () => {
    const args: { self?: unknown } = {};
    args.self = args;
    const card = approvalCard({ approvalId: 'ap-1', requesterId: 'assistant', tool: 'test', args });
    expect(JSON.stringify(card)).toContain('<unserializable args>');
  });

  it('includes the ellipsis in the approval argument length limit', () => {
    const card = approvalCard({ approvalId: 'ap', requesterId: 'alice', tool: 'test', args: 'x'.repeat(600) });
    expect(card.body?.[1]).toMatchObject({ facts: expect.arrayContaining([
      { title: 'Args', value: `"${'x'.repeat(498)}…` },
    ]) });
  });

  it('formats error and rejected results', () => {
    expect(formatToolResult({ status: 'error', error: 'boom' })).toMatch(/boom/);
    expect(formatToolResult({ status: 'rejected', message: 'no' })).toMatch(/no/);
  });
});

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

  it('truncates an over-long tool name', () => {
    const tool = 't'.repeat(200);
    const text = JSON.stringify(runningCard(tool));
    expect(text).toContain(`Running ${'t'.repeat(63)}…`);
    expect(text).not.toContain(tool);
  });
});
