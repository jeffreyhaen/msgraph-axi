# Changelog

All notable changes to this project are documented here. The format is based on
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project
follows [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Documentation & Infrastructure

- Run the workflows on `actions/checkout@v6`, `actions/setup-node@v6` and
  `pnpm/action-setup@v6`; the v4 actions still targeted Node 20, which GitHub runs on
  Node 24 with a deprecation warning.
- Move the `Releasing` steps under `Development` in `README.md`.

## [0.2.0] - 2026-09-13

### Changed (breaking)

- `mail send` no longer delivers mail on `--execute` alone. `--execute` now saves a
  draft and returns `sent: false, draft: true` with the message id; only
  `--send --execute` (or `mail send --draft <id> --execute` on a reviewed draft)
  delivers. Sending was irreversible, unreported and only gated by a flag that said
  nothing about delivery; callers that relied on the old behaviour must add `--send`.
- `mail send --attach` uploads the files with the draft through Graph instead of the
  backend send command (3 MB per message, the Graph limit).

### Added

- Calendar commands default to your mailbox time zone (Graph `mailboxSettings`,
  falling back to this machine's zone, then UTC) and echo the zone they used, so
  times are no longer silently read or written as UTC. `--timezone` still wins.
- `calendar create`/`update`: `--show-as free|tentative|busy|oof|workingElsewhere`
  sets the free/busy state of the event.
- `calendar availability`: a `legend` for the `availabilityView` codes and a
  `timezone` field for the window it queried.
- `calendar suggest`: the `x/y` available count includes the organizer, so `2/2`
  reads as "everyone free" rather than counting attendees only.
- `user search <term>`: directory lookup by display-name, first-name, last-name,
  mail or upn prefix (with `--limit`).

### Fixed

- Payloads now always travel as a file on Windows: cmd.exe also mangled an argument
  that mixes quotes with shell metacharacters (`<`, `>`, `&`, `|`), so an HTML body or a
  plain body containing `<` or `&` failed with an opaque error. `--body-type HTML`
  therefore works on Windows now. Multi-line bodies had the same cause: cmd.exe cut an
  argument at the first line break and caps the command line at 8191 characters, both
  without an error, so the backend refuses what it cannot carry instead of sending
  corrupt data.
- `calendar suggest`: report Graph's confidence as a percentage again (it was
  multiplied by 100, so `100` printed as `10000%`); the mock fixture now feeds
  real Graph values, which is why the wrong scale went unnoticed.
- `calendar agenda`: accept a bare local `--start`/`--end` instead of failing with
  the backend's "not a valid ISO date-time" error.
- Errors: only suggest `auth status` for authentication failures. A `403` now points
  at missing Graph permissions and a bad flag value at the input, instead of sending
  the caller to the connection.
- `raw --body @payload.json` hands the file to the backend instead of inlining it, so
  large and multi-line payloads survive on Windows.
- Payloads beyond the Windows command-line limit travel through a temp `@file`
  automatically, for every Graph bridge (mail draft, events, getSchedule,
  findMeetingTimes).

### Documentation & Infrastructure

- Document the time-zone precedence, the `--start`/`--end` formats, the
  `availabilityView` legend and the `suggest` versus `availability` difference in
  `README.md`, `SKILL.md` and `calendar --help`.
- `SKILL.md` spells out the draft-first mail flow and the rules agents must follow
  before delivering mail.

## [0.1.1] - 2026-09-13

### Fixed

- `calendar availability`: correctly parse Graph OData array payload returned by
  `getSchedule`.
- `mail list`: use `$top` query parameter through direct Graph request to properly
  respect `--limit`.
- `user get`: correct the user lookup endpoint URL.

### Documentation & Infrastructure

- Add interface comparison table and empirical token reduction benchmarks to README.
- Add Entra ID multi-tenant app registration guide with minimal permissions.
- Add repo header banner and badges for npm, CI, license, and TOON.
- Pin pnpm version 9 in GitHub Actions CI and release workflows.

## [0.1.0] - 2026-09-13

### Added

- TOON CLI over Microsoft Graph (Outlook mail + calendar), one backend: the CLI
  for Microsoft 365 (`m365`), dedicated commands plus the `m365 request` bridge.
- `auth status|login|logout` — sign-in state; device-code, browser, secret,
  certificate, password, identity and federated-identity logins.
- `mail list|read|send|delete` — messages with compact defaults, dry-run
  `--execute` gates and exact `--confirm <id>` on destructive ops.
- `mail search --search "<query>"` — Graph message search.
- `mail draft` + `mail send --draft <id>` — save a draft, then send it.
- `mail thread <conversationId>` — the whole conversation, oldest first.
- `mail attachment get <id> --message <id> [--out <path>]` — download files.
- `calendar list|agenda|create|update|cancel|delete` — calendars and events.
- `calendar availability --schedules` — free/busy via `getSchedule`.
- `calendar suggest --attendees` — meeting slot search via `findMeetingTimes`.
- `user get <upn>` — profile plus manager lookup.
- `raw <path>` — any Graph endpoint through `m365 request`.
- Home view (no args): sign-in status, unread count and today's agenda.
- AXI conventions: TOON on stdout, `{ error, code, help[] }` errors, exit codes
  0/1/2, unknown-flag rejection, `--limit`/`--fields`/`--full`.
- CI (build, typecheck, tests on Node 20/22, Ubuntu + Windows) and npm release
  workflow on version tags.