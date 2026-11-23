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
});
