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
                      [--mailbox <upn>] [--sender <upn>] [--send] [--execute]
msgraph-axi mail delete <id> [--user <upn>] [--execute --confirm <id>]
msgraph-axi mail search --search "<query>" [--user <upn>] [--limit 20]
msgraph-axi mail draft --to a@x.com --subject "..." --body "..." [--cc a] [--bcc b]
                        [--execute]
msgraph-axi mail send --draft <id> [--execute]
msgraph-axi mail thread <conversationId> [--user <upn>] [--full]
msgraph-axi mail attachment get <attachmentId> --message <messageId> [--out <path>]
```

`mail list` defaults to `id, from, subject, received`; `--fields` adds more columns
(e.g. `isRead,hasAttachments`). Cells truncate at 200 characters — `--full` lifts that.
`mail read` shows a snippet only; use `--full` precisely when the body is needed, and
pass its text back as-is.

`mail search` uses Graph `$search` (subject, body, sender). `mail thread`
returns the whole conversation oldest-first with snippets; `--full` adds bodies.
`mail attachment get` needs the message id from `mail list`/`mail read`; it writes to
`./<attachment-name>` unless `--out` is given.

### Sending mail

Mail is **never** delivered by `--execute` alone. `mail send --execute` saves a draft
(`sent: false, draft: true` plus the id) and `--send --execute` delivers that draft in
the same step, so there is always an inspectable message before anything leaves the
mailbox:

```sh
msgraph-axi mail draft --to a@x.com --subject "..." --body "..." --execute
msgraph-axi mail send --draft <id> --execute          # deliver a reviewed draft
msgraph-axi mail send --to a@x.com --subject "..." --body "..." --execute   # draft only
msgraph-axi mail send --to a@x.com --subject "..." --body "..." --send --execute
```

Rules for agents:

- Only deliver mail when the user explicitly asked for *sending*. "Maak een mail"
  means a draft: report the draft id and the exact recipients and let the user send it.
- Never invent recipients or addresses; resolve people with `user search <term>` and
  echo the resolved `upn` back to the user before drafting.
- Repeat the recipient, subject and a short body summary in your reply, so a wrong
  address or subject is caught before a human clicks send.
- `--execute` without `--send` is never a failure: report it as a saved draft, not as
  a sent mail.

`--body-type HTML` only when you actually pass HTML; plain text is the norm. Recipients
are comma-separated. `--attach` reads the files and uploads them with the draft (3 MB
per message, the Graph limit); larger files belong in a link, not an attachment.

A draft created through Graph carries no Outlook signature — signatures live in the
mail client, not in the mailbox, so nothing is added on the way out. Tell the user to
paste theirs in Outlook before sending, or include it in `--body` when they ask for it.

## Calendar

```sh
msgraph-axi calendar list [--user <upn>] [--limit 20]
msgraph-axi calendar agenda [--start <iso>] [--end <iso>] [--calendar <id|name>]
                            [--timezone <tz>] [--user <upn>] [--limit 20] [--fields a,b]
msgraph-axi calendar create --subject "..." --start <iso> --end <iso>
                            [--body "..."] [--location "..."] [--attendees a@x.com]
                            [--show-as free|tentative|busy|oof|workingElsewhere]
                            [--timezone <tz>] [--calendar <id>] [--execute]
msgraph-axi calendar update <id> [--subject "..."] [--start <iso>] [--end <iso>]
                               [--location "..."] [--body "..."] [--show-as <state>]
                               [--execute]
msgraph-axi calendar cancel <id> [--comment "..."] [--execute --confirm <id>]
msgraph-axi calendar delete <id> [--permanent] [--execute --confirm <id>]
msgraph-axi calendar availability --schedules a@x.com,b@y.com
                                  [--start <iso>] [--end <iso>] [--interval 30]
                                  [--timezone <tz>] [--full]
msgraph-axi calendar suggest --attendees a@x.com,b@y.com
                              [--duration 60] [--start <iso>] [--end <iso>]
                              [--candidates 5] [--timezone <tz>] [--full]
```

`calendar agenda` defaults to today through the next 7 days, sorted by start. Event
cells carry the time zone: `2026-03-15T12:00:00 CET`.

Every calendar command works in your mailbox time zone by default (Graph
`mailboxSettings`, falling back to this machine's zone, then UTC) and reports it in the
`timezone` field; pass `--timezone <tz>` to override. `--start`/`--end` accept a date
(`2026-03-15`), a wall clock in that zone (`2026-03-15T13:00:00`) or an explicit offset
(`2026-03-15T13:00:00+02:00`); anything else is rejected as a validation error, so a
bare `T00:00:00` never silently becomes UTC.

`calendar availability` maps to Graph `getSchedule` and summarizes each schedule as
`availabilityView` plus a `busy` count; the response carries a `legend` for those codes
(`0=free 1=tentative 2=busy 3=oof 4=workingElsewhere`) and `--full` expands the raw
schedule items. `calendar suggest` maps to Graph `findMeetingTimes` and returns
candidate slots with confidence (a percentage, as Graph reports it) and an `x/y`
available count that includes the organizer; `--full` breaks the availability down per
attendee.

`suggest` only knows the calendars Graph hands to `findMeetingTimes`: it misses shared
and delegated calendars that `availability` does see, so it can propose a slot where
you are already busy. Treat `suggest` as a hint and confirm the chosen slot with
`availability --schedules <you>,<attendee>` before creating the event with
`calendar create`.

`--show-as` sets how the event appears in free/busy: `free` keeps a 5-minute test or
reminder slot from blocking the room, `tentative`, `busy` (default), `oof` and
`workingElsewhere`.

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

## People

```sh
msgraph-axi user get <upn>
msgraph-axi user search <term> [--limit 15]
```

`user get` returns name, job title, department, office location, contact info and the
manager (`displayName <mail>`) — one call to answer "who is this person and who do they
report to?". `user search` finds people by display-name, first-name, last-name, mail or
upn **prefix** (one word: `alex`, `chen`, `a.chen`), which is the fastest way to turn
a first name from a meeting request into an upn for `calendar create --attendees`.
Reading other people's profiles and managers may require
`User.Read.All` or `Directory.Read.All` depending on the tenant; the signed-in
user only needs `User.Read`.

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