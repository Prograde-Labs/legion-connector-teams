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
