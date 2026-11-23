import { describe, expect, it } from 'vitest';
import { applyGroupPrefix } from './attribution.js';

describe('applyGroupPrefix', () => {
  it('prefixes the sender name by default', () => {
    expect(applyGroupPrefix('[{name}]: {message}', 'Alice', 'hello')).toBe('[Alice]: hello');
  });

  it('supports a custom template', () => {
    expect(applyGroupPrefix('{name} says: {message}', 'Alice', 'hello')).toBe('Alice says: hello');
  });

  it('leaves the message unchanged when attribution is disabled', () => {
    expect(applyGroupPrefix(null, 'Alice', 'hello')).toBe('hello');
  });

  it('works when the template has no name placeholder', () => {
    expect(applyGroupPrefix('{message}', 'Alice', 'hello')).toBe('hello');
  });

  it('does not interpret dollar patterns in the message', () => {
    expect(applyGroupPrefix('[{name}]: {message}', 'Alice', '$&')).toBe('[Alice]: $&');
    expect(applyGroupPrefix('[{name}]: {message}', 'Alice', '$$')).toBe('[Alice]: $$');
    expect(applyGroupPrefix('[{name}]: {message}', 'Alice', "$' $`")).toBe("[Alice]: $' $`");
  });

  it('does not interpret dollar patterns in the sender name', () => {
    expect(applyGroupPrefix('[{name}]: {message}', '$&', 'hello')).toBe('[$&]: hello');
  });

  it('does not rescan an injected sender name for placeholders', () => {
    expect(applyGroupPrefix('[{name}]: {message}', '{message}', 'hello')).toBe('[{message}]: hello');
  });

  it('replaces every message placeholder', () => {
    expect(applyGroupPrefix('{message} / {message}', 'Alice', 'hi')).toBe('hi / hi');
  });
});
