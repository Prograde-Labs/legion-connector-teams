#!/usr/bin/env node
// scripts/provision-teams-bot.mjs
import { execFileSync, spawnSync } from 'node:child_process';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { join, resolve } from 'node:path';

const COLOR_PNG =
  'iVBORw0KGgoAAAANSUhEUgAAAMAAAADACAYAAABS3GwHAAABiklEQVR42u3TMQ0AAAjAMGRiAP8JBsAEHz1qYMkiqwe+ChEwABgADAAGAAOAAcAAYAAwABgADAAGAAOAAcAAYAAwABgADAAGAAOAAcAAYAAwABgADAAGAAOAAcAAYAAwABgADAAGAAOAAcAAYAAwABgADAAGAAOAAcAAYAAwABgADIABRMAAYAAwABgADAAGAAOAAcAAYAAwABgADAAGAAOAAcAAYAAwABgADAAGAAOAAcAAYAAwABgADAAGAAOAAcAAYAAwABgADAAGAAOAAcAAYAAwABgADAAGAAOAAcAAYAAwAAYAA4ABwABgADAAGAAMAAYAA4ABwABgADAAGAAMAAYAA4ABwABgADAAGAAMAAYAA4ABwABgADAAGAAMAAYAA4ABwABgADAAGAAMAAYAA4ABwABgADAAGAAMAAYAA4ABMIAIGAAMAAYAA4ABwABgADAAGAAMAAYAA4ABwABgADAAGAAMAAYAA4ABwABgADAAGAAMAAYAA4ABwABgADAAGAAMAAaASwtC9faQjfST9wAAAABJRU5ErkJggg==';
const OUTLINE_PNG =
  'iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAAKUlEQVR42u3OIQEAAAACIP+f1hkWWEB6FgEBAQEBAQEBAQEBAQEBgXdgl/rw4tnPBf0AAAAASUVORK5CYII=';

const usage = 'Usage: provision-teams-bot.mjs --participant <id> --base-url <url>';
for (let i = 2; i < process.argv.length; i++) {
  if (['--participant', '--base-url', '--env', '--workspace', '--legion-url', '--legion-password'].includes(process.argv[i])) {
    const value = process.argv[i + 1];
    if (!value || value.startsWith('--')) {
      console.error(`${usage}\nMissing value for ${process.argv[i]}`);
      process.exit(2);
    }
    i++;
  }
}

function arg(name, fallback) {
  const index = process.argv.indexOf(`--${name}`);
  return index >= 0 ? process.argv[index + 1] : fallback;
}

const participantId = arg('participant');
const baseUrl = arg('base-url');
if (!participantId || !baseUrl) {
  console.error(usage);
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
  const zipCheck = spawnSync('zip', ['-v'], { stdio: 'ignore' });
  if (zipCheck.error || zipCheck.status !== 0) {
    throw new Error('zip is required; install the zip command before provisioning');
  }
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
    const match = line.match(/^\s*([A-Z0-9_]+)\s*=(.*)$/);
    if (match) {
      let value = match[2].trim();
      if (value.length >= 2 && ((value.startsWith('"') && value.endsWith('"')) || (value.startsWith("'") && value.endsWith("'")))) {
        value = value.slice(1, -1).trim();
      }
      env[match[1]] = value;
    }
  }
}
if (!dryRun) {
  const invalid = ['CLIENT_ID', 'CLIENT_SECRET', 'TENANT_ID'].filter((key) => {
    const value = env[key]?.trim();
    return !value || /^<.*>$/.test(value);
  });
  if (invalid.length) throw new Error(`Missing or invalid ${invalid.join(', ')} in ${envFile}`);
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

if (!dryRun) {
  try {
    const login = await fetch(`${legionUrl}/api/auth/login`, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ name: 'operator', password: legionPassword }),
    });
    const loginBody = (await login.json().catch(() => ({}))) ?? {};
    const { token } = loginBody;
    if (!login.ok || typeof token !== 'string' || !token.trim()) {
      throw new Error(`Legion login failed: ${loginBody.error ?? login.status}`);
    }
    const response = await fetch(`${legionUrl}/api/execute?stream=false`, {
      method: 'POST',
      headers: { authorization: `Bearer ${token.trim()}`, 'content-type': 'application/json' },
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
  } catch (error) {
    console.error(`Teams app ${appId} was created; credentials are in ${envFile}. Manually add the teams-bot identity and store the secret for this app ID, or delete the app in Teams before rerunning provisioning.`);
    throw error;
  }
}

const botIdentity = { connector: 'teams-bot', externalId: appId };
const identities = participant.identities ?? [];
if (!identities.some((i) => i.connector === botIdentity.connector && i.externalId === botIdentity.externalId)) {
  participant.identities = [...identities, botIdentity];
  await writeFile(participantPath, JSON.stringify(participant, null, 2));
}

if (!dryRun) {
  try {
    const link = execFileSync('teams', ['app', 'get', appId, '--install-link'], { encoding: 'utf8' }).trim();
    console.log(`Install: ${link}`);
  } catch {
    console.error('Could not retrieve install link; run this command manually:');
    console.log(`teams app get ${appId} --install-link`);
  }
}

console.log(JSON.stringify({ participantId, appId, endpoint, outDir }, null, 2));
