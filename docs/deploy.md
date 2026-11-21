# Deploying the Teams connector

## Build and mount

Build the connector with Node.js >=22.12.0 using `npm install` and `npm run build`. Mount the built `dist/` read-only into the Legion container. For the spec's module path below, mount it at `/app/legion-connector-teams/dist` (and make the connector's `package.json` and runtime `node_modules` available at `/app/legion-connector-teams`). Alternatively, if mounting `dist/` directly at `/app/legion-connector-teams`, use `/app/legion-connector-teams/index.js` as the module path and make its runtime dependencies resolvable there. The mount and module path must agree.

Add the following entry to the external connectors in `.legion/config.json`:

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

Persist the workspace and connector state directory across restarts. Configure `LEGION_SECRETS_KEY` in Legion's environment and keep it secure and stable; the connector requires Legion's SecretStore. Each bot's owning participant needs a `teams-bot` identity with its app ID and a SecretStore value under `teams-bot/<appId>` containing JSON `{ "appId": "...", "clientSecret": "...", "tenantId": "..." }`. Do not put secrets in config, participant files, or git.

## Transport

HTTP is the default. Route public HTTPS traffic for `teams.prograde-labs.com` through your reverse proxy or tunnel to `legion-collective:3978` (or the Legion host's exposed port 3978). Preserve the path: each bot's registered messaging endpoint is `https://teams.prograde-labs.com/teams/<participantId>/api/messages`. `GET /health` returns HTTP 200 without authentication. Ensure your container/network exposes the port to the proxy.

Alternatively set `"transport": "socket"`. Experimental Socket Mode needs no inbound port, public endpoint, or tunnel, and opens one outbound socket per configured geo for each bot. It is intended for new bots. The provisioning script still requires a `--base-url` argument.

## Tenant and provisioning

Use a Microsoft 365 tenant with Teams and custom-app upload/sideloading enabled; personal Teams accounts cannot sideload custom apps. A Microsoft 365 Business Basic trial expires after about a month, so durable operation requires a paid seat. Install the Teams Developer CLI and the external `zip` command on the provisioning host; `zip` may be absent on minimal hosts.

Use `scripts/provision-teams-bot.mjs` for each participant or `scripts/provision-all.mjs` for a selection. See the [README](../README.md#provisioning) for arguments, credentials, and the dry-run warning. Manually install every bot in Teams; registration alone does not install it for users. The directory is rescanned every 60 seconds by default.

## Operations and troubleshooting

- Watch for `[legion] external connector 'teams' started` in Legion logs. This only confirms connector startup, not that each bot is usable.
- Monitor per-bot startup and shutdown and `[teams]` initialization, start, stop, directory, and rebuild errors. The current bot pool logs per-bot failures, not successful per-bot start/stop events. HTTP reconciliation restarts the listener on rescans.
- Check for 403 delivery or conversation-creation errors, including an app not installed for the user. Install that bot manually and retry; outbound failures are logged as `Teams message delivery failed` or `Teams conversation creation failed`.
- Check for missing or malformed `teams-bot` secrets and app ID mismatches; such bots are skipped. Confirm the participant identity and SecretStore key match.
- Socket Mode has no HTTP health endpoint. For HTTP, check `/health` separately from end-to-end Teams delivery.
