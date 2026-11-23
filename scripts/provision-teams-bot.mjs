#!/usr/bin/env node
// scripts/provision-teams-bot.mjs
import { execFileSync, spawnSync } from 'node:child_process';
import { chmod, mkdir, readFile, writeFile } from 'node:fs/promises';
import { join, resolve } from 'node:path';

const COLOR_PNG =
  'iVBORw0KGgoAAAANSUhEUgAAAMAAAADACAYAAABS3GwHAAABiklEQVR42u3TMQ0AAAjAMGRiAP8JBsAEHz1qYMkiqwe+ChEwABgADAAGAAOAAcAAYAAwABgADAAGAAOAAcAAYAAwABgADAAGAAOAAcAAYAAwABgADAAGAAOAAcAAYAAwABgADAAGAAOAAcAAYAAwABgADAAGAAOAAcAAYAAwABgADIABRMAAYAAwABgADAAGAAOAAcAAYAAwABgADAAGAAOAAcAAYAAwABgADAAGAAOAAcAAYAAwABgADAAGAAOAAcAAYAAwABgADAAGAAOAAcAAYAAwABgADAAGAAOAAcAAYAAwAAYAA4ABwABgADAAGAAMAAYAA4ABwABgADAAGAAMAAYAA4ABwABgADAAGAAMAAYAA4ABwABgADAAGAAMAAYAA4ABwABgADAAGAAMAAYAA4ABwABgADAAGAAMAAYAA4ABMIAIGAAMAAYAA4ABwABgADAAGAAMAAYAA4ABwABgADAAGAAMAAYAA4ABwABgADAAGAAMAAYAA4ABwABgADAAGAAMAAaASwtC9faQjfST9wAAAABJRU5ErkJggg==';
const OUTLINE_PNG =
  'iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAAKUlEQVR42u3OIQEAAAACIP+f1hkWWEB6FgEBAQEBAQEBAQEBAQEBgXdgl/rw4tnPBf0AAAAASUVORK5CYII=';

const usage = 'Usage: provision-teams-bot.mjs --participant <id> [--base-url <url> | --socket] [--force] [--dry-run]';
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
const socket = process.argv.includes('--socket');
if (!participantId || (!baseUrl && !socket)) {
  console.error(usage);
  process.exit(2);
}

const workspace = resolve(arg('workspace', process.env.LEGION_WORKSPACE ?? process.cwd()));
const envFile = resolve(arg('env', join(workspace, '.env.teams')));
const legionUrl = arg('legion-url', 'http://127.0.0.1:3000');
const legionPassword = arg('legion-password', process.env.LEGION_BOOTSTRAP_PASSWORD);
const dryRun = process.argv.includes('--dry-run');
const force = process.argv.includes('--force');

const participantPath = join(workspace, '.legion/collective/participants', `${participantId}.json`);
const participant = JSON.parse(await readFile(participantPath, 'utf8'));
const previous = participant.identities?.find((identity) => identity.connector === 'teams-bot');
const reuse = Boolean(previous && !force);
const endpoint = baseUrl ? `${baseUrl.replace(/\/$/, '')}/teams/${participantId}/api/messages` : null;

if (reuse) {
  if (!dryRun && (!previous.externalId?.trim() || /^<.*>$/.test(previous.externalId))) {
    throw new Error('Existing teams-bot identity is a placeholder; use --force to replace it.');
  }
  console.log(`Reusing Teams app ${previous.externalId}; skipping teams app create.`);
} else {
  console.log(`${dryRun ? 'Would create' : 'Creating'} a Teams app${previous ? ' to replace the existing identity (--force)' : ''}.`);
}

if (!dryRun) {
  const zipCheck = spawnSync('zip', ['-v'], { stdio: 'ignore' });
  if (zipCheck.error || zipCheck.status !== 0) {
    throw new Error('zip is required; install the zip command before provisioning');
  }
}

async function protectEnvFile() {
  await chmod(envFile, 0o600).catch((error) => {
    if (error.code !== 'ENOENT') throw error;
  });
}

const env = {};
if (!dryRun && !reuse) {
  await protectEnvFile();
  const mask = process.umask(0o077);
  try {
    execFileSync('teams', socket
      ? ['-y', 'app', 'create', '--name', participant.name, '--messaging-mode', 'socket', '--env', envFile]
      : ['-y', 'app', 'create', '--name', participant.name, '--endpoint', endpoint, '--env', envFile],
    { stdio: ['inherit', 'ignore', 'ignore'] });
  } catch {
    throw new Error(`teams app create failed; check Teams and ${envFile} before retrying.`);
  } finally {
    process.umask(mask);
    await protectEnvFile();
  }
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
  const invalid = ['CLIENT_ID', 'CLIENT_SECRET', 'TENANT_ID'].filter((key) => {
    const value = env[key]?.trim();
    return !value || /^<.*>$/.test(value);
  });
  if (invalid.length) throw new Error(`Missing or invalid ${invalid.join(', ')} in ${envFile}`);
}
const appId = reuse ? previous.externalId : env.CLIENT_ID ?? '<app-id>';

const manifest = {
  $schema: 'https://developer.microsoft.com/en-us/json-schemas/teams/v1.19/MicrosoftTeams.schema.json',
  manifestVersion: '1.19',
  version: '1.0.0',
  id: appId,
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
  bots: [{ botId: appId, scopes: ['personal', 'groupChat'], supportsFiles: false, isNotificationOnly: false }],
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

async function execute(token, tool, args) {
  const response = await fetch(`${legionUrl}/api/execute?stream=false`, {
    method: 'POST',
    headers: { authorization: `Bearer ${token}`, 'content-type': 'application/json' },
    body: JSON.stringify({ tool, args }),
  });
  if (!response.ok) throw new Error(`${tool} failed (HTTP ${response.status})`);
  const { result } = (await response.json().catch(() => ({}))) ?? {};
  if (result?.status !== 'success') throw new Error(`${tool} did not succeed`);
}

const botIdentity = { connector: 'teams-bot', externalId: appId };
async function replaceBotIdentity(current) {
  const identities = current.identities ?? [];
  current.identities = [...identities.filter((identity) => identity.connector !== 'teams-bot'), botIdentity];
  await writeFile(participantPath, JSON.stringify(current, null, 2));
}

if (dryRun) {
  if (!reuse) await replaceBotIdentity(participant);
} else {
  try {
    const login = await fetch(`${legionUrl}/api/auth/login`, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ name: 'operator', password: legionPassword }),
    });
    const loginBody = (await login.json().catch(() => ({}))) ?? {};
    const { token } = loginBody;
    if (!login.ok || typeof token !== 'string' || !token.trim()) {
      throw new Error(`Legion login failed (HTTP ${login.status})`);
    }
    if (!reuse) {
      await execute(token.trim(), 'set_secret', {
        participantId,
        key: `teams-bot/${appId}`,
        secret: JSON.stringify({ appId, clientSecret: env.CLIENT_SECRET, tenantId: env.TENANT_ID }),
      });
    }
    await execute(token.trim(), 'pair', { participantId, ...botIdentity });
  } catch (error) {
    if (!reuse) console.error(`Teams app ${appId} was created; credentials are in ${envFile}. Store the secret and pair this app ID manually, or delete the app in Teams before retrying.`);
    throw error;
  }
  if (force && previous) {
    // The core pair tool only merges identities. Preserve any other changes it persisted.
    await replaceBotIdentity(JSON.parse(await readFile(participantPath, 'utf8')));
    console.log('Replaced the teams-bot identity in the participant file. Restart Legion or reconcile its collective to retire the old identity.');
  }
  try {
    const link = execFileSync('teams', ['app', 'get', appId, '--install-link'], { encoding: 'utf8', stdio: ['ignore', 'pipe', 'ignore'] }).trim();
    console.log(`Install: ${link}`);
  } catch {
    console.error('Could not retrieve install link; run this command manually:');
    console.log(`teams app get ${appId} --install-link`);
  }
}

console.log(JSON.stringify({ participantId, appId, endpoint, outDir }, null, 2));
