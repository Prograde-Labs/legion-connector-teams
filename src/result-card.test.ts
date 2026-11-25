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
