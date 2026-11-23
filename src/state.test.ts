import { chmod, mkdir, mkdtemp, readdir, readFile, stat, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { describe, expect, it, vi } from 'vitest';
import { emptyState, StateStore } from './state.js';

const isRoot = typeof process.getuid === 'function' && process.getuid() === 0;

async function backups(dir: string): Promise<string[]> {
  return (await readdir(dir)).filter((name) => name.startsWith('state.json.bak-'));
}

describe('StateStore', () => {
  it('returns empty v2 state when the file is missing', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'teams-state-'));
    await expect(new StateStore(dir).load()).resolves.toEqual(emptyState());
  });

  it('round-trips v2 state', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'teams-state-'));
    const store = new StateStore(dir);
    const state = emptyState();
    state.conversations['app:alice'] = {
      conversationId: 'conv-1',
      conversationType: 'personal',
      externalId: 'aad-1',
      updatedAt: '2026-11-21T00:00:00.000Z',
    };
    state.conversations['app:group:19:g@thread.v2'] = {
      conversationId: '19:g@thread.v2',
      conversationType: 'groupChat',
      externalId: '19:g@thread.v2',
      name: 'Engineering',
      updatedAt: '2026-11-21T00:00:00.000Z',
    };
    state.legionConversationTargets['thread-1'] = 'app:group:19:g@thread.v2';
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

  it('serializes overlapping saves', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'teams-state-'));
    const store = new StateStore(dir);
    const first = emptyState();
    first.legionConversations['a'] = 'legion-a';
    const second = emptyState();
    second.legionConversations['b'] = 'legion-b';

    await expect(Promise.all([store.save(first), store.save(second)])).resolves.toHaveLength(2);

    await expect(store.load()).resolves.toEqual(second);
  });

  it('migrates v1 conversations and pending pairs', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'teams-state-'));
    await writeFile(join(dir, 'state.json'), JSON.stringify({
      version: 1,
      conversations: {
        'app:alice': {
          conversationId: 'conv-1',
          tenantId: 'tenant-1',
          serviceUrl: 'https://smba.trafficmanager.net/teams/',
          aadObjectId: 'aad-1',
          displayName: 'Alice',
          updatedAt: '2026-11-21T00:00:00.000Z',
        },
      },
      pendingPairs: {
        ABC234: {
          appId: 'app',
          aadObjectId: 'aad-2',
          expiresAt: '2026-11-22T00:00:00.000Z',
        },
      },
      legionConversations: { 'app:alice': 'thread-1' },
    }));
    await expect(new StateStore(dir).load()).resolves.toEqual({
      version: 2,
      conversations: {
        'app:alice': {
          conversationId: 'conv-1',
          conversationType: 'personal',
          externalId: 'aad-1',
          tenantId: 'tenant-1',
          serviceUrl: 'https://smba.trafficmanager.net/teams/',
          displayName: 'Alice',
          updatedAt: '2026-11-21T00:00:00.000Z',
        },
      },
      pendingPairs: {
        ABC234: {
          appId: 'app',
          connector: 'teams',
          externalId: 'aad-2',
          expiresAt: '2026-11-22T00:00:00.000Z',
        },
      },
      legionConversations: { 'app:alice': 'thread-1' },
      legionConversationTargets: {},
    });
  });

  it('defaults a missing target map on load', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'teams-state-'));
    await writeFile(join(dir, 'state.json'), JSON.stringify({
      version: 2,
      conversations: {},
      pendingPairs: {},
      legionConversations: {},
    }));
    await expect(new StateStore(dir).load()).resolves.toEqual(emptyState());
  });

  it('backs up corrupt state and starts fresh', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'teams-state-'));
    await writeFile(join(dir, 'state.json'), '{not json');
    const store = new StateStore(dir);
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    try {
      await expect(store.load()).resolves.toEqual(emptyState());
      const [saved] = await backups(dir);
      expect(warn).toHaveBeenCalledWith(`[teams] Corrupt state backed up to ${join(dir, saved)}`);
      expect(JSON.stringify(warn.mock.calls)).not.toContain('{not json');
    } finally { warn.mockRestore(); }
    const raw = await readFile(join(dir, 'state.json'), 'utf8').catch(() => null);
    expect(raw).toBeNull();
    const saved = await backups(dir);
    expect(saved).toHaveLength(1);
    await expect(readFile(join(dir, saved[0]), 'utf8')).resolves.toBe('{not json');
  });

  it.each([
    [
      'non-object map',
      '{"version":2,"conversations":[],"pendingPairs":{},"legionConversations":{},"legionConversationTargets":{}}',
    ],
    [
      'unsupported version',
      '{"version":3,"conversations":{},"pendingPairs":{},"legionConversations":{}}',
    ],
  ])('backs up valid JSON with malformed shape: %s', async (_label, contents) => {
    const dir = await mkdtemp(join(tmpdir(), 'teams-state-'));
    await writeFile(join(dir, 'state.json'), contents);
    const store = new StateStore(dir);
    await expect(store.load()).resolves.toEqual(emptyState());
    const raw = await readFile(join(dir, 'state.json'), 'utf8').catch(() => null);
    expect(raw).toBeNull();
    const saved = await backups(dir);
    expect(saved).toHaveLength(1);
    await expect(readFile(join(dir, saved[0]), 'utf8')).resolves.toBe(contents);
  });

  it('propagates read errors other than ENOENT', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'teams-state-'));
    await mkdir(join(dir, 'state.json'));
    await expect(new StateStore(dir).load()).rejects.toMatchObject({ code: 'EISDIR' });
  });

  it.skipIf(isRoot)('propagates backup failures', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'teams-state-'));
    await writeFile(join(dir, 'state.json'), '{not json');
    await chmod(dir, 0o500);
    try {
      await expect(new StateStore(dir).load()).rejects.toMatchObject({ code: 'EACCES' });
    } finally {
      await chmod(dir, 0o700);
    }
  });
});
