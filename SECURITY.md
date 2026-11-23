# Security policy

## Reporting a vulnerability

Report security issues privately through GitHub's **Report a vulnerability** button on the
repository's Security tab. Please do not open a public issue for a suspected vulnerability.

Include the affected commit or version, reproduction steps, and the impact you observed. We
aim to acknowledge reports within a few business days.

## Secrets

This connector handles Microsoft Teams bot credentials. Keep them in Legion's SecretStore
under `teams-bot/<appId>`, and never commit `.env.teams`, participant files containing
secrets, or connector state. The provisioning script creates `.env.teams` with mode `0600`;
see the README for details.
