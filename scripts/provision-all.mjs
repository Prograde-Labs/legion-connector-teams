#!/usr/bin/env node
// scripts/provision-all.mjs
import { execFileSync } from 'node:child_process';
import { readFile, readdir } from 'node:fs/promises';
import { join, resolve } from 'node:path';

const workspace = resolve(process.env.LEGION_WORKSPACE ?? process.cwd());
const baseUrl = process.argv[process.argv.indexOf('--base-url') + 1];
if (!baseUrl) {
  console.error('Usage: provision-all.mjs --base-url <url>');
  process.exit(2);
}

const dir = join(workspace, '.legion/collective/participants');
for (const file of await readdir(dir)) {
  if (!file.endsWith('.json')) continue;
  const participant = JSON.parse(await readFile(join(dir, file), 'utf8'));
  if (!participant.identities?.some((i) => i.connector === 'teams-bot')) continue;
  execFileSync('node', [
    join(import.meta.dirname, 'provision-teams-bot.mjs'),
    '--participant', participant.id,
    '--base-url', baseUrl,
    '--workspace', workspace,
  ], { stdio: 'inherit' });
}
