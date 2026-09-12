---
name: msgraph-axi
description: Use msgraph-axi to work with Microsoft Graph — Outlook mail, calendars, events, meeting availability, and raw Graph endpoints — through token-efficient TOON output. Use when the task involves Microsoft Graph, Outlook mail, messages, calendars, events, meetings, free/busy, getSchedule, or sending mail for a Microsoft 365 account.
user-invocable: false
---

# msgraph-axi

Agent-ergonomic TOON wrapper for Microsoft Graph with Outlook mail and calendar as the
first surface. It runs the [CLI for Microsoft 365](https://pnp.github.io/cli-microsoft365/)
(`m365`) as its single backend: dedicated `outlook` commands where they exist, and the
generic `m365 request` bridge for the rest. The m365 syntax is a backend detail — never
expose `m365 ...` command lines to the user.

Install globally: `npm install -g @jeffreyhaen/msgraph-axi @pnp/cli-microsoft365`. If
`msgraph-axi` is not on PATH, prefix with `npx -y @jeffreyhaen/msgraph-axi`.

## Orientation

Run `msgraph-axi` with no arguments first. It prints the signed-in account plus the
unread inbox count and today's first events — enough to act without a second call.

```sh
msgraph-axi                    # dashboard (sign-in + unread + today)
msgraph-axi auth status        # signed-in account and tenant
msgraph-axi auth login         # interactive device-code sign-in (only interactive step)
```

## Mail

```sh
msgraph-axi mail list [--folder inbox|drafts|<name|id>] [--start <iso>] [--end <iso>]
                      [--user <upn>] [--limit 20] [--fields a,b] [--full]
msgraph-axi mail read <id> [--user <upn>] [--full]     # --full adds the body content
msgraph-axi mail send --to a@x.com,b@y.com --subject "..." --body "..."
                      [--cc a@x.com] [--bcc b@x.com] [--body-type text|HTML]
                      [--importance low|normal|high] [--attach path1,path2]
                      [--mailbox <upn>] [--sender <upn>] [--execute]
msgraph-axi mail delete <id> [--user <upn>] [--execute --confirm <id>]
```

`mail list` defaults to `id, from, subject, received`; `--fields` adds more columns
(e.g. `isRead,hasAttachments`). Cells truncate at 200 characters — `--full` lifts that.
`mail read` shows a snippet only; use `--full` precisely when the body is needed, and
pass its text back as-is.

### Sending mail

`mail send` dry-runs by default: it prints the recipient preview and asks for
`--execute`. Only run it when the user asked to send. `--body-type HTML` is the default
only if you pass HTML; plain text is the norm. Recipients are comma-separated.

## Calendar

```sh
msgraph-axi calendar list [--user <upn>] [--limit 20]
msgraph-axi calendar agenda [--start <iso>] [--end <iso>] [--calendar <id|name>]
                            [--timezone <tz>] [--user <upn>] [--limit 20] [--fields a,b]
msgraph-axi calendar create --subject "..." --start <iso> --end <iso>
                            [--body "..."] [--location "..."] [--attendees a@x.com]
                            [--timezone <tz>] [--calendar <id>] [--execute]
msgraph-axi calendar update <id> [--subject "..."] [--start <iso>] [--end <iso>]
                               [--location "..."] [--body "..."] [--execute]
msgraph-axi calendar cancel <id> [--comment "..."] [--execute --confirm <id>]
msgraph-axi calendar delete <id> [--permanent] [--execute --confirm <id>]
msgraph-axi calendar availability --schedules a@x.com,b@y.com
                                  [--start <iso>] [--end <iso>] [--interval 30]
                                  [--timezone <tz>] [--full]
```

`calendar agenda` defaults to today through the next 7 days, sorted by start. Event
cells carry the time zone: `2026-03-15T12:00:00 CET`. `calendar availability` maps to
Graph `getSchedule` and summarizes each schedule as `availabilityView` plus a `busy`
count; `--full` expands the raw schedule items.

### Meetings and write gates

`create` and `update` are dry-runs without `--execute`. `cancel` and `delete` require
both `--execute` and `--confirm <id>` — the id must match exactly. `cancel` sends
cancellation to attendees; `delete --permanent` empties the deleted-items bin. Only run
these when the user asked — and never cancel or delete an event the user only asked to
move or reschedule: use `calendar update` with new `--start`/`--end` instead.

`--calendar <id>` on create/update targets a specific calendar (take the id from
`calendar list`). `--user <upn>` works on every command; the default is the signed-in
account. When a user pastes a meeting request or an event snippet, extract subject,
start/end, location and attendees from it and pass them explicitly.

## Escape hatch

Anything without a dedicated command goes through `m365 request`:

```sh
msgraph-axi raw me/mailFolders/inbox?$select=unreadItemCount
msgraph-axi raw me/messages --method post --body '{"subject":"Hi","body":{"content":"..."}}'
msgraph-axi raw me/events/AAMk... --method patch --body @payload.json
msgraph-axi raw me/calendar/getSchedule --method post \
  --body '{"schedules":["a@x.com"],"startTime":{"dateTime":"2026-03-15T09:00:00","timeZone":"UTC"},"endTime":{"dateTime":"2026-03-15T18:00:00","timeZone":"UTC"}}'
```

Paths are relative to `https://graph.microsoft.com/v1.0/`. GET is the default; write
methods (`post`, `put`, `patch`, `delete`) require `--execute`. `--body @<file>` reads
the file; `--query 'a=b&c=d'` appends query parameters.

## Conventions

- Output is TOON on stdout; errors are TOON too, with a `help` list of next steps.
- Exit codes: 0 success (including no-ops), 1 runtime error, 2 usage error.
- Unknown flags are rejected by name — read the `help` line and retry once.
- Lists take `--limit` and `--fields a,b`; detail views truncate and take `--full`.
- `--user <upn>` selects a mailbox explicitly; the default is the signed-in account.
- Mutating commands (`mail send`, `mail delete`, `calendar create/update/cancel/delete`)
  change real state — only run them when the user asked, and respect the `--execute` /
  `--confirm <id>` gates instead of inventing bypasses.
- `msgraph-axi auth login` is the only interactive step (device code flow); every other
  command is fully non-interactive.
- On Windows/Git Bash, do not pass large or multiline content through a `.cmd` shim as
  an interpolated CLI argument; write it to a file and use `--body @file` instead.
  File paths starting with `/` need `MSYS_NO_PATHCONV=1` in front of the command.
- The CLI for Microsoft 365 is a backend detail: never quote `m365 ...` syntax back to
  the user, and never suggest installing PnP-specific flags.