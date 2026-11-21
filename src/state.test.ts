import { mkdtemp, readFile, stat, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { emptyState, StateStore } from './state.js';

describe('StateStore', () => {
  it('returns empty state when the file is missing', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'teams-state-'));
    await expect(new StateStore(dir).load()).resolves.toEqual(emptyState());
  });

  it('round-trips state', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'teams-state-'));
    const store = new StateStore(dir);
    const state = emptyState();
    state.conversations['app:alice'] = {
      conversationId: 'conv-1',
      updatedAt: '2026-11-21T00:00:00.000Z',
    };
    await store.save(state);
    await expect(store.load()).resolves.toEqual(state);
  });

  it('writes the file with mode 0600', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'teams-state-'));
    const store = new StateStore(dir);
    await store.save(emptyState());
    const mode = (await stat(join(dir, 'state.json'))).mode & 0o777;
    expect(mode).toBe(0o600);
  });

  it('backs up corrupt state and starts fresh', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'teams-state-'));
    await writeFile(join(dir, 'state.json'), '{not json');
    const store = new StateStore(dir);
    await expect(store.load()).resolves.toEqual(emptyState());
    const raw = await readFile(join(dir, 'state.json'), 'utf8').catch(() => null);
    expect(raw).toBeNull();
  });
});
