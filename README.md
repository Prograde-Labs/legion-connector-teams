# Legion Microsoft Teams connector

The Teams connector exposes Legion participants as individual Teams bots for personal 1:1 conversations. It supports text, conversation continuity, pairing, and approval cards. Messages outside personal chats are rejected, and messages, installs, and approval actions must report the tenant configured in the bot's credentials.

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
| `provisionerId` | `operator` | Caller for auto-provisioning `create_user` or `modify_user`. |
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

Users have a `teams` identity whose `externalId` is their Entra `aadObjectId` (the connector falls back to the Teams sender ID if absent). Unknown users default to receiving a pairing code. They give the code to a registered user, who runs `/pair <code> <participant-id>` in a personal chat with the same bot. The connector calls Legion's core `pair` tool, subject to the caller's AuthEngine policy. Configure `pairingParticipantId` if another participant should make that call. Auto-provisioning requires `create_user` (or `modify_user` for an existing account) for `provisionerId` and `pair` for the pairing caller. It grants the provisioned user `communicate` and `approval_response`.

Fallback maps unknown users to one shared Legion participant (`fallbackUserId`) but keeps a separate conversation and Legion thread per Teams sender. Proactive delivery is best effort: an exact participant route takes precedence; otherwise the bot sends to the most recently active fallback chat. It cannot reliably target a particular fallback sender, so avoid using shared fallback for private proactive messages.

## Chat commands

- `/new` — start a new Legion conversation.
- `/pair <code> <participant-id>` — pair a pending Teams identity with a Legion participant.
- `/whoami` — show identity and conversation information.
- `/help` — list commands.

## Provisioning

Run from the connector checkout, with `--workspace` pointing to the Legion workspace containing the participant files. Replace these example values with your own:

```sh
# HTTP transport (public endpoint)
node scripts/provision-teams-bot.mjs --participant assistant --base-url https://teams.prograde-labs.com --workspace /srv/legion
node scripts/provision-all.mjs --base-url https://teams.prograde-labs.com --workspace /srv/legion --participants assistant,operator

# Socket Mode (no public endpoint)
node scripts/provision-teams-bot.mjs --participant assistant --socket --workspace /srv/legion
```

The script reuses an existing `teams-bot` app ID and its stored credentials without calling `teams app create`. For a new bot it registers the app and stores credentials via Legion's `set_secret` tool. The credentials file (`--env`, default `<workspace>/.env.teams`) is restricted to mode `0600`. The script builds the app package with `zip` and calls Legion's `pair` tool as `operator`, so a running connector discovers new identities on its next rescan. Set `--legion-url` (default `http://127.0.0.1:3000`) and `--legion-password` (default `LEGION_BOOTSTRAP_PASSWORD`) as needed.

Use `--force` to register a new bot and replace the identity in the participant file. The running collective still has the old identity until it is reconciled or Legion is restarted; restart promptly before making further changes to that participant. The old Teams registration is not deleted. Both scripts accept `--force`, `--dry-run`, and `--socket`; without `--participants` the batch script selects non-retired agents. **Dry runs still write manifests and icons and set a placeholder identity for bots they would create**, so use a disposable workspace. With `--socket`, `teams app create` runs with `--messaging-mode socket`; `--base-url` is optional.

Install each generated app manually in Teams using the printed install link or package. If SecretStore setup fails after app creation, credentials remain in the environment file; follow the script's recovery message before rerunning to avoid duplicate registrations.

## Testing

```sh
npm test
```

A live Socket Mode integration test is available as opt-in with `LEGION_TEAMS_INTEGRATION=1` and `TEAMS_CLIENT_ID`, `TEAMS_CLIENT_SECRET`, and `TEAMS_TENANT_ID`. Do not put those credentials in source control.

See [deployment](docs/deploy.md) for container, proxy, and operational notes.
