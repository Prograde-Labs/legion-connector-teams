# Legion Microsoft Teams connector

The Teams connector exposes [Legion](https://github.com/Prograde-Labs/legion) participants as individual Teams bots for personal 1:1 chats and group chats. It supports text, conversation continuity, pairing, and approval cards. Team channels are not supported yet. Messages, installs, and approval actions must report the tenant configured in the bot's credentials.

## Prerequisites and installation

- Node.js >=22.12.0. The Teams SDK packages are pinned to 2.2.1; the HTTP server uses Express 5.
- A Microsoft 365 tenant with Teams and custom-app upload/sideloading enabled. Personal Teams accounts cannot sideload custom apps. A Business Basic trial expires after about a month; durable use requires a paid seat.
- For HTTP, a public HTTPS endpoint routed to the connector; alternatively, experimental Socket Mode needs no public endpoint.
- For provisioning, the Teams Developer CLI (`teams`), the external `zip` command (not installed on every host), and a running Legion instance with SecretStore enabled.
- A sibling checkout of the [Legion](https://github.com/Prograde-Labs/legion) repository. `@legion-collective/core` and `@legion-collective/types` are resolved from `../legion/packages/*` and are not yet published to npm.

```sh
# Clone Legion next to this repository and build its core packages first.
git clone https://github.com/Prograde-Labs/legion.git ../legion
(cd ../legion && npm ci && npm run build)

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
| `groupChats` | `true` | Enable group-chat participation. |
| `groupSenderPrefix` | `"[{name}]: {message}"` | Sender prefix for group messages; `null` disables it and the template must contain `{message}`. |
| `provisionerId` | `operator` | Caller for auto-provisioning `create_user` or `modify_user`. |
| `pairingParticipantId` | unset | Caller for `pair`; defaults to the sender, or the provisioner during auto-provisioning. |
| `pairingTtlSeconds` | `600` | Pairing code lifetime. |
| `maxDetail` | `standard` | `minimal`, `standard`, or `full`. |
| `toolDisplay` | `status` | `off`, `status`, or `cards`. |
| `aiLabel` | `true` | Label outgoing messages as AI-generated. |
| `feedback` | `false` | Reserved for future use. |
| `typingIntervalMs` | `4000` | Typing interval; `0` disables it. |
| `maxMessageChars` | `20000` | Outbound message chunk size. |
| `learnOutboundThread` | `true` | Record an outbound-initiated Legion conversation as the chat's current thread, so the Teams reply continues it. |

HTTP listens on port 3978 by default, with each bot at `/teams/<participantId>/api/messages` and an unauthenticated `GET /health`. Socket Mode (`"transport": "socket"`) is experimental, intended for new bots, and opens one socket per configured geo per bot; it has no HTTP listener. Teams streaming is supported in 1:1 and group chats; team channels are not supported. Teams allows one open stream per chat, so inbound turns are serialized per conversation.

## Identity and credentials

Each Teams-facing agent has a `teams-bot` identity whose `externalId` is its app ID. Its Legion SecretStore entry is scoped to that participant, keyed `teams-bot/<appId>`, and contains JSON `{ "appId": "...", "clientSecret": "...", "tenantId": "..." }`. Set `LEGION_SECRETS_KEY` for Legion; the connector cannot start without its SecretStore. Never commit credentials or the provisioning environment file.

Users have a `teams` identity whose `externalId` is their Entra `aadObjectId` (the connector falls back to the Teams sender ID if absent). Unknown users default to receiving a pairing code. They give the code to a registered user, who runs `/pair <code> <participant-id>` in a personal chat or a group chat with the same bot. The connector calls Legion's core `pair` tool, subject to the caller's AuthEngine policy. Configure `pairingParticipantId` if another participant should make that call. Auto-provisioning requires `create_user` (or `modify_user` for an existing account) for `provisionerId` and `pair` for the pairing caller. It grants the provisioned user `communicate` and `approval_response`.

Group chats are modelled as one Legion `user` participant with a `teams-group` identity whose `externalId` is the Teams conversation ID. The group uses the same `unknownUserPolicy` as personal chats. With `auto-provision`, the participant ID is `teams-group-<name>-<hash>`; unnamed groups use `Teams group <hash>`. Group messages are prefixed with the sender's name by default. `/new` and `/whoami` act on the group, while `/pair` runs as the registered human who typed it. Replies are routed back through the Legion conversation target map, so a fallback participant paired to several groups still answers in the originating group.

Fallback maps unknown users to one shared Legion participant (`fallbackUserId`) but keeps a separate conversation and Legion thread per Teams sender, and a separate conversation per group, so fallback conversations may now be groups as well as personal chats. Proactive delivery is best effort: the Legion conversation target map takes precedence, then an exact participant route; otherwise the bot sends to the most recently active personal fallback chat (group records are excluded from that scan). It cannot reliably target a particular fallback sender, so avoid using shared fallback for private proactive messages. Proactive failures are no longer silent: failed sends are logged; both `failed` and `not-reachable` are returned to the initiating `communicate` call with a sanitized reason, and thread learning for outbound-initiated conversations is on by default (`learnOutboundThread: true`). An unresolved group cannot run `/pair` from inside itself; its pairing code must be redeemed from another chat with the same bot.

## Outbound initiation

A participant that owns a `teams-bot` identity can start a conversation with an
existing Teams participant using `communicate({ to, message })`. The target must
already exist in the collective with a `teams` (personal) or `teams-group`
(group chat) identity. The connector reports `delivered`, `not-reachable`, or
`failed`; `accepted` is part of the core contract but Teams never emits it. A
failed send makes the synchronous `communicate` call return an error with a
sanitized reason.

On a successful send the connector records the Legion conversation as the chat's
current thread by default (`learnOutboundThread: true`), so the Teams user's next
reply continues that conversation. Set `learnOutboundThread: false` to keep the
previous thread; the conversation target map is still updated. Cold outreach,
rosters, retries, and delivery receipts are not supported yet.

## Chat commands

- `/new` — start a new Legion conversation for the current chat (the group, in a group chat).
- `/pair <code> <participant-id>` — pair a pending identity; in a group chat it runs as the registered human who typed it.
- `/whoami` — show the current participant, chat, sender, and bot app.
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

## Security

Bot client secrets live in Legion's SecretStore under `teams-bot/<appId>`. The provisioning environment file (`<workspace>/.env.teams`) is local, mode `0600`, and must never be committed. Report vulnerabilities privately as described in [SECURITY.md](SECURITY.md).

## License

Apache License 2.0. See [LICENSE](LICENSE).
