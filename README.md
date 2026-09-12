# msgraph-axi

Agent-ergonomic [TOON](https://axi.md/) wrapper for Microsoft Graph — Outlook mail and
calendar first, with a raw Graph bridge for everything else. Built on the
[CLI for Microsoft 365](https://pnp.github.io/cli-microsoft365/) (`m365`) as a single
backend, following the [AXI](https://github.com/kunchenguid/axi) design principles.

## Install

```sh
npm install -g @jeffreyhaen/msgraph-axi @pnp/cli-microsoft365
msgraph-axi auth login
msgraph-axi          # dashboard
```

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