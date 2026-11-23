import { describe, expect, it } from 'vitest';
import { SerialQueue } from './serialize.js';

describe('SerialQueue', () => {
  it('serializes tasks for the same key', async () => {
    const queue = new SerialQueue();
    const order: string[] = [];
    let release!: () => void;
    const gate = new Promise<void>((resolve) => { release = resolve; });
    const first = queue.run('chat-1', async () => {
      order.push('first-start');
      await gate;
      order.push('first-end');
    });
    const second = queue.run('chat-1', async () => { order.push('second'); });

    await Promise.resolve();
    expect(order).toEqual(['first-start']);
    release();
    await Promise.all([first, second]);
    expect(order).toEqual(['first-start', 'first-end', 'second']);
  });

  it('runs tasks for different keys concurrently', async () => {
    const queue = new SerialQueue();
    const order: string[] = [];
    let release!: () => void;
    const gate = new Promise<void>((resolve) => { release = resolve; });
    const first = queue.run('chat-1', async () => {
      order.push('first-start');
      await gate;
      order.push('first-end');
    });
    const second = queue.run('chat-2', async () => { order.push('second'); });

    await Promise.resolve();
    expect(order).toEqual(['first-start', 'second']);
    release();
    await Promise.all([first, second]);
    expect(order).toEqual(['first-start', 'second', 'first-end']);
  });

  it('continues after a rejected task', async () => {
    const queue = new SerialQueue();
    await expect(queue.run('chat-1', async () => { throw new Error('boom'); })).rejects.toThrow('boom');
    await expect(queue.run('chat-1', async () => undefined)).resolves.toBeUndefined();
  });
});
