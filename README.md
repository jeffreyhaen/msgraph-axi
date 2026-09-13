# msgraph-axi (Microsoft Graph axi)

[![ci](https://github.com/jeffreyhaen/msgraph-axi/actions/workflows/ci.yml/badge.svg)](https://github.com/jeffreyhaen/msgraph-axi/actions/workflows/ci.yml)
[![npm](https://img.shields.io/npm/v/%40jeffreyhaen%2Fmsgraph-axi.svg)](https://www.npmjs.com/package/@jeffreyhaen/msgraph-axi)
[![license](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![node](https://img.shields.io/badge/node-%3E%3D20-brightgreen.svg)](package.json)

<p align="center">
  <img src="assets/msgraph-axi-header.png" alt="msgraph-axi header">
</p>

Agent-ergonomic [TOON](https://toonformat.dev/) wrapper for **Microsoft Graph** — Outlook mail and
calendar first, with a raw Graph bridge for everything else. Built on the
[CLI for Microsoft 365](https://pnp.github.io/cli-microsoft365/) (`m365`) as a single
backend, following the [AXI](https://github.com/kunchenguid/axi) design principles.

## Why not an MCP server

A Microsoft Graph MCP server loads its full schema into the agent's context window on every
turn (~25k-40k tokens before any work is done). A skill-based AXI costs ~55 tokens until the
agent actually needs it, outputs compact [TOON](https://toonformat.dev/) instead of bloated
raw Graph JSON payloads, and enforces dry-run safety gates on mutating actions.

## Install & Setup

Install globally (recommended for repeated use):

```sh
npm install -g @jeffreyhaen/msgraph-axi @pnp/cli-microsoft365
```

For a one-off invocation without installing:

```sh
npx -y @jeffreyhaen/msgraph-axi --help
```

## Agent integration

Install the skill globally so an agent loads the usage guide on demand:

```sh
npx skills add jeffreyhaen/msgraph-axi --skill msgraph-axi -g
```

Omit `-g` to install the skill for the current project only.

### Initial Entra ID / M365 Setup

The underlying CLI for Microsoft 365 requires an Entra ID application registration in your tenant:

1. **Automatic setup (quickest):**
   Run `m365 setup` and choose **Create a new app registration**. It registers an app with the CLI presets and stores the App ID in your local config.

2. **Existing or manual App Registration:**
   If your organization already has an Entra app registered for CLI / developer use (or you create one manually):
   - Set redirect URI (Public client/mobile & desktop): `https://login.microsoftonline.com/common/oauth2/nativeclient`
   - Grant **Delegated** Microsoft Graph permissions: `Mail.ReadWrite`, `Mail.Send`, `Calendars.ReadWrite`, `User.Read`
   - Save the App ID and Tenant in your CLI config:
     ```sh
     m365 cli config set --key clientId --value "<your-app-id>"
     m365 cli config set --key tenantId --value "<your-tenant-id>"
     ```

Then log in:
```sh
msgraph-axi auth login
msgraph-axi          # dashboard
```

See [docs/authentication.md](docs/authentication.md) for headless / CI flows and details.

## Commands

| Command | Purpose |
|---|---|
| `auth status\|login\|logout` | m365 sign-in state; device-code and headless logins (see [Authentication](docs/authentication.md)) |
| `mail list [--folder] [--start] [--end]` | messages with compact defaults |
| `mail read <id> [--full]` | snippet by default, body with `--full` |
| `mail send ... [--execute]` | dry-run unless `--execute`; `--draft <id>` sends a saved draft |
| `mail delete <id> [--execute --confirm <id>]` | destructive, double-gated |
| `mail search --search "..."` | Graph message search |
| `mail draft --to ... [--execute]` | save a draft, review, then `mail send --draft` |
| `mail thread <conversationId>` | the whole conversation, oldest first |
| `mail attachment get <id> --message <id> [--out <path>]` | download an attachment |
| `calendar list` | calendars of the signed-in (or `--user`) account |
| `calendar agenda [--start] [--end] [--calendar]` | events, default today..+7d |
| `calendar create \| update [--execute]` | events via the Graph REST bridge |
| `calendar cancel \| delete [--execute --confirm <id>]` | destructive, double-gated |
| `calendar availability --schedules a@x.com` | free/busy via `getSchedule` |
| `calendar suggest --attendees a@x.com --duration 60` | meeting slot search via `findMeetingTimes` |
| `user get <upn>` | profile plus manager lookup |
| `raw <path> [--method] [--body] [--execute]` | any Graph endpoint through `m365 request` |

## Conventions

- TOON on stdout, structured `{ error, code, help[] }` errors on stdout.
- Exit codes: 0 success (incl. no-ops), 1 runtime error, 2 usage error.
- Unknown flags rejected by name; lists take `--limit` / `--fields a,b`; detail views
  truncate at 200 chars and lift with `--full`.
- Read-only by default; writes require `--execute`, destructive writes also
  `--confirm <id>`.

## Development

```sh
pnpm install
pnpm test          # vitest against a fake m365 fixture
pnpm typecheck
pnpm build
pnpm dev -- mail list   # run against a real m365 login
```

Tests run against `test/fixtures/m365-cli.js` — no Microsoft 365 tenant required. Live
smoke tests need `m365 login`; the fixture can be pointed at with
`MSGRAPH_AXI_M365_BIN`.

- [AXI — agent eXperience interface](https://axi.md/) · [kunchenguid/axi](https://github.com/kunchenguid/axi)
- [CLI for Microsoft 365](https://pnp.github.io/cli-microsoft365/) — the backend
- [Microsoft Graph](https://learn.microsoft.com/en-us/graph/overview) — the platform

## See also

- [docs/authentication.md](docs/authentication.md) — every login flow with concrete commands