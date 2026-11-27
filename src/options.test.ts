import { describe, expect, it } from 'vitest';
import { resolveOptions } from './options.js';

describe('resolveOptions group chat options', () => {
  it('defaults group chats on with the default attribution template', () => {
    const options = resolveOptions({ stateDir: '/tmp/x' });
    expect(options.groupChats).toBe(true);
    expect(options.groupSenderPrefix).toBe('[{name}]: {message}');
  });

  it('accepts a null attribution template', () => {
    expect(resolveOptions({ stateDir: '/tmp/x', groupSenderPrefix: null }).groupSenderPrefix).toBeNull();
  });

  it('accepts a custom attribution template', () => {
    expect(resolveOptions({ stateDir: '/tmp/x', groupSenderPrefix: '{message} — {name}' }).groupSenderPrefix)
      .toBe('{message} — {name}');
  });

  it('rejects a template without a message placeholder', () => {
    expect(() => resolveOptions({ stateDir: '/tmp/x', groupSenderPrefix: '{name}' }))
      .toThrow(/\{message\}/);
  });

  it('rejects an empty attribution template', () => {
    expect(() => resolveOptions({ stateDir: '/tmp/x', groupSenderPrefix: '' }))
      .toThrow(/\{message\}/);
  });

  it('can disable group chats', () => {
    expect(resolveOptions({ stateDir: '/tmp/x', groupChats: false }).groupChats).toBe(false);
  });
});

describe('resolveOptions outbound thread learning', () => {
  it('defaults learnOutboundThread on', () => {
    expect(resolveOptions({ stateDir: '/tmp/x' }).learnOutboundThread).toBe(true);
  });

  it('can disable learnOutboundThread', () => {
    expect(resolveOptions({ stateDir: '/tmp/x', learnOutboundThread: false }).learnOutboundThread)
      .toBe(false);
  });
});

describe('resolveOptions tools', () => {
  it('enables all tools by default', () => {
    const options = resolveOptions({ stateDir: '/tmp/x' });
    expect([...options.tools].sort()).toEqual(['cards', 'list', 'reactions']);
    expect(options.maxMessageRecords).toBe(2000);
    expect(options.maxCardBytes).toBe(28000);
  });

  it('accepts a subset', () => {
    expect([...resolveOptions({ stateDir: '/tmp/x', tools: ['cards'] }).tools]).toEqual(['cards']);
  });

  it('accepts false to disable every tool', () => {
    expect(resolveOptions({ stateDir: '/tmp/x', tools: false }).tools.size).toBe(0);
  });

  it('rejects an unknown tool name', () => {
    expect(() => resolveOptions({ stateDir: '/tmp/x', tools: ['nope' as never] })).toThrow(
      /unknown Teams tool/,
    );
  });

  it('rejects non-positive message and card limits', () => {
    expect(() => resolveOptions({ stateDir: '/tmp/x', maxMessageRecords: 0 })).toThrow(
      /maxMessageRecords/,
    );
    expect(() => resolveOptions({ stateDir: '/tmp/x', maxCardBytes: 0 })).toThrow(/maxCardBytes/);
  });
});

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
