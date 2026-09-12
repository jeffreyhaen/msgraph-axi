# Changelog

All notable changes to this project are documented here. The format is based on
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project
follows [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.1.0] - unreleased

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