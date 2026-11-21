#!/usr/bin/env node
// scripts/provision-all.mjs
import { execFileSync } from 'node:child_process';
import { readFile, readdir } from 'node:fs/promises';
import { join, resolve } from 'node:path';

function arg(name, fallback) {
  const index = process.argv.indexOf(`--${name}`);
  const value = index >= 0 ? process.argv[index + 1] : undefined;
  return value && !value.startsWith('--') ? value : fallback;
}

const workspace = resolve(arg('workspace', process.env.LEGION_WORKSPACE ?? process.cwd()));
const baseUrl = arg('base-url');
if (!baseUrl) {
  console.error('Usage: provision-all.mjs --base-url <url>');
  process.exit(2);
}

const dir = join(workspace, '.legion/collective/participants');
const selected = arg('participants')?.split(',');
const passthrough = [];
for (const name of ['env', 'legion-url', 'legion-password']) {
  const value = arg(name);
  if (value !== undefined) passthrough.push(`--${name}`, value);
}
if (process.argv.includes('--dry-run')) passthrough.push('--dry-run');
for (const file of await readdir(dir)) {
  if (!file.endsWith('.json')) continue;
  const participant = JSON.parse(await readFile(join(dir, file), 'utf8'));
  if (selected ? !selected.includes(participant.id) : participant.type !== 'agent' || participant.status === 'retired') continue;
  execFileSync('node', [
    join(import.meta.dirname, 'provision-teams-bot.mjs'),
    '--participant', participant.id,
    '--base-url', baseUrl,
    '--workspace', workspace,
    ...passthrough,
  ], { stdio: 'inherit' });
}
