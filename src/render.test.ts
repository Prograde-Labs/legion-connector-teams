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
