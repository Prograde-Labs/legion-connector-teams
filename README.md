# Legion Microsoft Teams connector

The Teams connector exposes Legion participants as individual Teams bots for personal 1:1 conversations. It supports text, conversation continuity, pairing, and approval cards; group chats and channels are not supported.

## Prerequisites and installation

- Node.js >=22.12.0. The Teams SDK packages are pinned to 2.2.1; the HTTP server uses Express 5.
- A Microsoft 365 tenant with Teams and custom-app upload/sideloading enabled. Personal Teams accounts cannot sideload custom apps. A Business Basic trial expires after about a month; durable use requires a paid seat.
- For HTTP, a public HTTPS endpoint routed to the connector; alternatively, experimental Socket Mode needs no public endpoint.
- For provisioning, the Teams Developer CLI (`teams`), the external `zip` command (not installed on every host), and a running Legion instance with SecretStore enabled.

```sh
npm install
npm run build
```

## Configuration

Add this connector entry to Legion's `.legion/config.json`, adjusting paths and tenant values for your deployment:

```json
{
  "name": "teams",
  "module": "/app/legion-connector-teams/dist/index.js",
  "enabled": true,
  "options": {
    "transport": "http",
    "port": 3978,
    "publicBaseUrl": "https://teams.prograde-labs.com",
    "tenantId": "<tenant-id>",
    "stateDir": "/workspace/.legion/connectors/teams",
    "unknownUserPolicy": "pairing",
    "maxDetail": "standard",
    "toolDisplay": "status",
    "aiLabel": true
  }
}
```

Options and defaults:

| Option | Default | Purpose |
| --- | --- | --- |
| `transport` | `http` | `http` or experimental `socket`. |
| `port` | `3978` | HTTP listener port. |
| `publicBaseUrl` | unset | Public URL for provisioning/configuration; does not configure the listener. |
| `tenantId` | unset | Tenant used when creating outbound conversations. |
| `stateDir` | `$LEGION_WORKSPACE/.legion/connectors/teams` | Connector state directory; current working directory replaces `$LEGION_WORKSPACE` when unset. |
| `rescanIntervalMs` | `60000` | Directory rescan interval. |
| `socketGeos` | `["amer", "emea", "apac"]` | Socket Mode regions. |
| `unknownUserPolicy` | `pairing` | `pairing`, `auto-provision`, `reject`, or `fallback`. |
| `fallbackUserId` | unset | Required for `fallback`. |
| `provisionerId` | `operator` | Caller for auto-provisioning `create_user`. |
| `pairingParticipantId` | unset | Caller for `pair`; defaults to the sender, or the provisioner during auto-provisioning. |
| `pairingTtlSeconds` | `600` | Pairing code lifetime. |
| `maxDetail` | `standard` | `minimal`, `standard`, or `full`. |
| `toolDisplay` | `status` | `off`, `status`, or `cards`. |
| `aiLabel` | `true` | Label outgoing messages as AI-generated. |
| `feedback` | `false` | Reserved for future use. |
| `typingIntervalMs` | `4000` | Typing interval; `0` disables it. |
| `maxMessageChars` | `20000` | Outbound message chunk size. |

HTTP listens on port 3978 by default, with each bot at `/teams/<participantId>/api/messages` and an unauthenticated `GET /health`. Socket Mode (`"transport": "socket"`) is experimental, intended for new bots, and opens one socket per configured geo per bot; it has no HTTP listener. Teams streaming is limited to 1:1 chats and a window of approximately two minutes.

## Identity and credentials

Each Teams-facing agent has a `teams-bot` identity whose `externalId` is its app ID. Its Legion SecretStore entry is scoped to that participant, keyed `teams-bot/<appId>`, and contains JSON `{ "appId": "...", "clientSecret": "...", "tenantId": "..." }`. Set `LEGION_SECRETS_KEY` for Legion; the connector cannot start without its SecretStore. Never commit credentials or the provisioning environment file.

Users have a `teams` identity whose `externalId` is their Entra `aadObjectId` (the connector falls back to the Teams sender ID if absent). Unknown users default to receiving a pairing code. An authorized user can run `/pair <code> <participant-id>`: the connector calls Legion's core `pair` tool, subject to the caller's AuthEngine policy. Configure `pairingParticipantId` if another participant should make that call. Auto-provisioning also requires `create_user` permission for `provisionerId` and `pair` permission for the pairing caller. Fallback maps unknown users to `fallbackUserId`; use it only when that trust model is appropriate.

## Chat commands

- `/new` — start a new Legion conversation.
- `/pair <code> <participant-id>` — pair a pending Teams identity with a Legion participant.
- `/whoami` — show identity and conversation information.
- `/help` — list commands.

## Provisioning

Run from the connector checkout, with `--workspace` pointing to the Legion workspace containing the participant files:

```sh
node scripts/provision-teams-bot.mjs --participant <id> --base-url https://teams.prograde-labs.com --workspace <workspace>
node scripts/provision-all.mjs --base-url https://teams.prograde-labs.com --workspace <workspace> --participants <id1,id2>
```

The first script registers a bot, creates a Teams app package, stores its credentials through Legion's `set_secret` API, and updates the participant identity. It uses `--env` (default `<workspace>/.env.teams`), `--legion-url` (default `http://127.0.0.1:3000`), and `--legion-password` (default `LEGION_BOOTSTRAP_PASSWORD`) when supplied. The bulk script passes these options through; without `--participants`, it selects non-retired agents. Both support `--dry-run`, but **dry-run still writes a manifest, icons, and a placeholder bot identity to the participant file**; use only a disposable workspace for it. Provisioning still requires `--base-url` even if Socket Mode will be used.

Install each generated app manually in Teams using the printed install link or package. If SecretStore setup fails after app creation, credentials remain in the environment file; follow the script's recovery message before rerunning to avoid duplicate registrations.

## Testing

```sh
npm test
```

A live Socket Mode integration test is planned as opt-in with `LEGION_TEAMS_INTEGRATION=1` and `TEAMS_CLIENT_ID`, `TEAMS_CLIENT_SECRET`, and `TEAMS_TENANT_ID`; it is not present in this checkout yet. Do not put those credentials in source control.

See [deployment](docs/deploy.md) for container, proxy, and operational notes.
