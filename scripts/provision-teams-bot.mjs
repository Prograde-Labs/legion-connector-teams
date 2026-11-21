#!/usr/bin/env node
// scripts/provision-teams-bot.mjs
import { execFileSync } from 'node:child_process';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { join, resolve } from 'node:path';

const COLOR_PNG =
  'iVBORw0KGgoAAAANSUhEUgAAAMAAAADACAYAAABS3GwHAAABiklEQVR42u3TMQ0AAAjAMGRiAP8JBsAEHz1qYMkiqwe+ChEwABgADAAGAAOAAcAAYAAwABgADAAGAAOAAcAAYAAwABgADAAGAAOAAcAAYAAwABgADAAGAAOAAcAAYAAwABgADAAGAAOAAcAAYAAwABgADAAGAAOAAcAAYAAwABgADIABRMAAYAAwABgADAAGAAOAAcAAYAAwABgADAAGAAOAAcAAYAAwABgADAAGAAOAAcAAYAAwABgADAAGAAOAAcAAYAAwABgADAAGAAOAAcAAYAAwABgADAAGAAOAAcAAYAAwAAYAA4ABwABgADAAGAAMAAYAA4ABwABgADAAGAAMAAYAA4ABwABgADAAGAAMAAYAA4ABwABgADAAGAAMAAYAA4ABwABgADAAGAAMAAYAA4ABwABgADAAGAAMAAYAA4ABMIAIGAAMAAYAA4ABwABgADAAGAAMAAYAA4ABwABgADAAGAAMAAYAA4ABwABgADAAGAAMAAYAA4ABwABgADAAGAAMAAaASwtC9faQjfST9wAAAABJRU5ErkJggg==';
const OUTLINE_PNG =
  'iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAAKUlEQVR42u3OIQEAAAACIP+f1hkWWEB6FgEBAQEBAQEBAQEBAQEBgXdgl/rw4tnPBf0AAAAASUVORK5CYII=';

function arg(name, fallback) {
  const index = process.argv.indexOf(`--${name}`);
  return index >= 0 ? process.argv[index + 1] : fallback;
}

const participantId = arg('participant');
const baseUrl = arg('base-url');
if (!participantId || !baseUrl) {
  console.error('Usage: provision-teams-bot.mjs --participant <id> --base-url <url>');
  process.exit(2);
}

const workspace = resolve(arg('workspace', process.env.LEGION_WORKSPACE ?? process.cwd()));
const envFile = resolve(arg('env', join(workspace, '.env.teams')));
const legionUrl = arg('legion-url', 'http://127.0.0.1:3000');
const legionPassword = arg('legion-password', process.env.LEGION_BOOTSTRAP_PASSWORD);
const dryRun = process.argv.includes('--dry-run');

const participantPath = join(workspace, '.legion/collective/participants', `${participantId}.json`);
const participant = JSON.parse(await readFile(participantPath, 'utf8'));
const endpoint = `${baseUrl.replace(/\/$/, '')}/teams/${participantId}/api/messages`;

if (!dryRun) {
  execFileSync('teams', [
    'app', 'create',
    '--name', participant.name,
    '--endpoint', endpoint,
    '--env', envFile,
  ], { stdio: 'inherit' });
}

const env = {};
if (!dryRun) {
  for (const line of (await readFile(envFile, 'utf8')).split(/\r?\n/)) {
    const match = line.match(/^([A-Z0-9_]+)=(.*)$/);
    if (match) env[match[1]] = match[2];
  }
}
const appId = env.CLIENT_ID ?? '<app-id>';
const clientSecret = env.CLIENT_SECRET ?? '<client-secret>';
const tenantId = env.TENANT_ID ?? '<tenant-id>';

const manifest = {
  $schema: 'https://developer.microsoft.com/en-us/json-schemas/teams/v1.19/MicrosoftTeams.schema.json',
  manifestVersion: '1.19',
  version: '1.0.0',
  id: appId,
  packageName: `com.prograde-labs.legion.${participantId}`,
  developer: {
    name: 'Prograde Labs',
    websiteUrl: 'https://prograde-labs.com',
    privacyUrl: 'https://prograde-labs.com',
    termsOfUseUrl: 'https://prograde-labs.com',
  },
  icons: { color: 'color.png', outline: 'outline.png' },
  name: { short: participant.name, full: participant.name },
  description: {
    short: `Legion agent ${participant.name}`,
    full: `Legion agent ${participant.name} on Microsoft Teams.`,
  },
  accentColor: '#5865F2',
  bots: [{ botId: appId, scopes: ['personal'], supportsFiles: false, isNotificationOnly: false }],
  permissions: ['identity', 'messageTeamMembers'],
  validDomains: [],
};

const outDir = join(workspace, '.legion/teams-apps', participantId);
await mkdir(outDir, { recursive: true });
await writeFile(join(outDir, 'manifest.json'), JSON.stringify(manifest, null, 2));
await writeFile(join(outDir, 'color.png'), Buffer.from(COLOR_PNG, 'base64'));
await writeFile(join(outDir, 'outline.png'), Buffer.from(OUTLINE_PNG, 'base64'));

if (!dryRun) {
  execFileSync('zip', ['-j', join(outDir, `${participantId}.zip`), 'manifest.json', 'color.png', 'outline.png'], { cwd: outDir, stdio: 'inherit' });
}

const botIdentity = { connector: 'teams-bot', externalId: appId };
const identities = participant.identities ?? [];
if (!identities.some((i) => i.connector === botIdentity.connector && i.externalId === botIdentity.externalId)) {
  participant.identities = [...identities, botIdentity];
  await writeFile(participantPath, JSON.stringify(participant, null, 2));
}

if (!dryRun) {
  const login = await fetch(`${legionUrl}/api/auth/login`, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ name: 'operator', password: legionPassword }),
  });
  const { token } = await login.json();
  const response = await fetch(`${legionUrl}/api/execute?stream=false`, {
    method: 'POST',
    headers: { authorization: `Bearer ${token}`, 'content-type': 'application/json' },
    body: JSON.stringify({
      tool: 'set_secret',
      args: {
        participantId,
        key: `teams-bot/${appId}`,
        secret: JSON.stringify({ appId, clientSecret, tenantId }),
      },
    }),
  });
  if (!response.ok) throw new Error(`set_secret failed: ${response.status}`);
  const { result } = (await response.json());
  if (result?.status !== 'success') {
    throw new Error(`set_secret failed: ${result?.error ?? result?.status ?? 'unknown'}`);
  }
  console.log(`Install: teams app get ${appId} --install-link`);
}

console.log(JSON.stringify({ participantId, appId, endpoint, outDir }, null, 2));
