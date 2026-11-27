import { execFile } from 'node:child_process';
import { mkdir, mkdtemp, readFile, rm, stat, writeFile } from 'node:fs/promises';
import { createServer } from 'node:http';
import type { AddressInfo } from 'node:net';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { promisify } from 'node:util';
import { expect, it } from 'vitest';

const exec = promisify(execFile);
const botIdentity = (externalId: string) => ({ connector: 'teams-bot', externalId });
const userIdentity = { connector: 'teams', externalId: 'aad-1' };

it.each(['new', 'reuse', 'force'] as const)('provisions a %s bot without exposing credentials', async (mode) => {
  const workspace = await mkdtemp(join(tmpdir(), 'teams-provision-'));
  const participantDir = join(workspace, '.legion/collective/participants');
  const participantPath = join(participantDir, 'assistant.json');
  const envFile = join(workspace, '.env.teams');
  const participant = { id: 'assistant', name: 'Assistant', type: 'agent', identities: [
    userIdentity, ...(mode === 'new' ? [] : [botIdentity('existing-app')]),
  ] };
  const requests: unknown[] = [];
  const readParticipant = async () => JSON.parse(await readFile(participantPath, 'utf8')) as typeof participant;
  const server = createServer(async (req, res) => {
    try {
      let raw = '';
      for await (const chunk of req) raw += chunk;
      const body = JSON.parse(raw);
      requests.push({ method: req.method, url: req.url, authorization: req.headers.authorization, body });
      if (req.url === '/api/auth/login') {
        res.end(JSON.stringify({ token: 'operator-token' }));
        return;
      }
      if (body.tool === 'pair') {
        // Simulate the core tool's persistence: append rather than replace identities.
        const current = await readParticipant();
        if (!current.identities.some((id) => id.connector === 'teams-bot' && id.externalId === body.args.externalId)) {
          current.identities.push(botIdentity(body.args.externalId));
          await writeFile(participantPath, JSON.stringify(current));
        }
      }
      res.end(JSON.stringify({ result: { status: 'success' } }));
    } catch {
      res.writeHead(500).end();
    }
  });
  try {
    await mkdir(participantDir, { recursive: true });
    await writeFile(participantPath, JSON.stringify(participant));
    await writeFile(join(workspace, 'zip'), '#!/bin/sh\nexit 0\n', { mode: 0o700 });
    await writeFile(join(workspace, 'teams'), `#!/bin/sh
env_file=""
app_id=""
prev=""
for arg in "$@"; do
  if [ "$prev" = "--env" ]; then env_file="$arg"; fi
  if [ "$prev" = "get" ]; then app_id="$arg"; fi
  prev="$arg"
done
case " $* " in
  *" create "*)
    : > '${join(workspace, 'created')}'
    printf '%s\n' 'CLIENT_ID=new-app' 'CLIENT_SECRET=stub-secret' 'TENANT_ID=tenant' > "$env_file"
    echo stub-secret
    ;;
  *" get "*)
    echo "https://example.test/install/$app_id"
    ;;
  *) exit 1 ;;
esac
`, { mode: 0o700 });
    await new Promise<void>((resolve) => server.listen(0, '127.0.0.1', resolve));
    const { port } = server.address() as AddressInfo;
    const { stdout, stderr } = await exec(process.execPath, [
      'scripts/provision-teams-bot.mjs', '--participant', 'assistant', '--workspace', workspace,
      '--base-url', 'https://example.test', '--legion-url', `http://127.0.0.1:${port}`,
      ...(mode === 'force' ? ['--force'] : []),
    ], { env: { ...process.env, PATH: `${workspace}:${process.env.PATH}`, LEGION_BOOTSTRAP_PASSWORD: 'stub-password' } });
    const appId = mode === 'reuse' ? 'existing-app' : 'new-app';
    expect(stdout).toContain(`https://example.test/install/${appId}`);
    expect(stdout + stderr).not.toMatch(/stub-secret|operator-token|stub-password/);
    const manifest = JSON.parse(await readFile(join(workspace, '.legion/teams-apps/assistant/manifest.json'), 'utf8'));
    expect(manifest.bots).toEqual([expect.objectContaining({ botId: appId, scopes: ['personal', 'groupChat'] })]);
    expect(manifest.version).toBe('1.1.0');
    expect(manifest.commandLists).toEqual([
      {
        scopes: ['personal', 'groupChat'],
        commands: [
          { title: 'new', description: 'Start a new conversation' },
          { title: 'switch', description: 'List and switch to an existing conversation' },
          { title: 'compact', description: 'Compact this conversation' },
          { title: 'context', description: 'Show usage and context details' },
          { title: 'whoami', description: 'Show identity and chat details' },
          { title: 'pair', description: 'Pair a pending Teams identity' },
          { title: 'help', description: 'List commands' },
        ],
      },
    ]);
    expect((await readParticipant()).identities).toEqual([userIdentity, botIdentity(appId)]);
    const request = (tool: string, args: object) => ({ method: 'POST', url: '/api/execute?stream=false',
      authorization: 'Bearer operator-token', body: { tool, args } });
    expect(requests).toEqual([
      { method: 'POST', url: '/api/auth/login', authorization: undefined, body: { name: 'operator', password: 'stub-password' } },
      ...(mode === 'reuse' ? [] : [request('set_secret', { participantId: 'assistant', key: 'teams-bot/new-app',
        secret: JSON.stringify({ appId, clientSecret: 'stub-secret', tenantId: 'tenant' }) })]),
      request('pair', { participantId: 'assistant', ...botIdentity(appId) }),
    ]);
    if (mode === 'reuse') {
      expect(stdout).toContain('skipping teams app create');
      await expect(stat(join(workspace, 'created'))).rejects.toMatchObject({ code: 'ENOENT' });
    } else {
      expect((await stat(envFile)).mode & 0o777).toBe(0o600);
    }
    if (mode === 'force') expect(stdout).toContain('Restart Legion or reconcile its collective');
  } finally {
    await new Promise<void>((resolve) => server.close(() => resolve()));
    await rm(workspace, { recursive: true, force: true });
  }
});
