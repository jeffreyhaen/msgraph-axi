import { runAxiCli, AxiError, type AxiCliCommand } from "axi-sdk-js";
import { PnpCliBackend } from "./backend.js";
import { VERSION } from "./version.js";
import { authStatus, authLogin, authLogout, type AuthContext } from "./commands/auth.js";
import {
  mailList,
  mailRead,
  mailSend,
  mailDelete,
  type MailContext,
} from "./commands/mail.js";
import {
  calendarList,
  calendarAgenda,
  calendarCreate,
  calendarUpdate,
  calendarCancel,
  calendarDelete,
  calendarAvailability,
  calendarSuggest,
  type CalendarContext,
} from "./commands/calendar.js";
import {
  mailSearch,
  mailDraft,
  mailThread,
  mailAttachmentGet,
  type MailFlowContext,
} from "./commands/mailflow.js";
import { userGet, userSearch, type UserContext } from "./commands/user.js";
import { raw, type RawContext } from "./commands/raw.js";
import { home, type HomeContext } from "./commands/home.js";

type AppContext = AuthContext &
  MailContext &
  CalendarContext &
  MailFlowContext &
  RawContext &
  UserContext &
  HomeContext;

const GLOBAL_FLAGS = "--user <upn>  --limit N  --fields a,b  --full  --execute  --help";

/** Adapt a handler with a narrower context to the app-wide context. */
function cmd(handler: AxiCliCommand<any>): AxiCliCommand<AppContext> {
  return handler as unknown as AxiCliCommand<AppContext>;
}

function sub(
  handlers: Record<string, AxiCliCommand<AppContext>>,
  help: string,
): AxiCliCommand<AppContext> {
  return async (args, context) => {
    const [subcommand, ...rest] = args;
    if (subcommand === undefined || subcommand === "--help") {
      throw new AxiError(
        "A subcommand is required",
        "VALIDATION_ERROR",
        help.split("\n").filter(Boolean),
      );
    }
    const handler = handlers[subcommand] as AxiCliCommand<AppContext> | undefined;
    if (!handler) {
      throw new AxiError(
        `Unknown subcommand: ${subcommand}`,
        "VALIDATION_ERROR",
        help.split("\n").filter(Boolean),
      );
    }
    return handler(rest, context);
  };
}

const commands: Record<string, AxiCliCommand<AppContext>> = {
  "auth": sub(
    {
      status: cmd(authStatus),
      login: cmd(authLogin),
      logout: cmd(authLogout),
    },
    "Valid subcommands: status, login, logout",
  ),
  "mail": sub(
    {
      list: cmd(mailList),
      read: cmd(mailRead),
      send: cmd(mailSend),
      delete: cmd(mailDelete),
      search: cmd(mailSearch),
      draft: cmd(mailDraft),
      thread: cmd(mailThread),
      attachment: sub(
        { get: cmd(mailAttachmentGet) },
        "Valid: mail attachment get <attachmentId> --message <messageId> [--out <path>]",
      ),
    },
    [
      "Valid subcommands:",
      "  mail list [--folder <name|id>] [--start <iso>] [--end <iso>]",
      "  mail read <id> [--full]",
      "  mail send --to <a,b> --subject \"...\" --body \"...\" [--execute]",
      "  mail send --draft <id> [--execute]",
      "  mail delete <id> [--execute --confirm <id>]",
      "  mail search --search \"<query>\"",
      "  mail draft --to <a> --subject \"...\" --body \"...\" [--execute]",
      "  mail thread <conversationId>",
      "  mail attachment get <attachmentId> --message <messageId> [--out <path>]",
    ].join("\n"),
  ),
  "calendar": sub(
    {
      list: cmd(calendarList),
      agenda: cmd(calendarAgenda),
      create: cmd(calendarCreate),
      update: cmd(calendarUpdate),
      cancel: cmd(calendarCancel),
      delete: cmd(calendarDelete),
      availability: cmd(calendarAvailability),
      suggest: cmd(calendarSuggest),
    },
    [
      "Valid subcommands:",
      "  calendar list",
      "  calendar agenda [--start <iso>] [--end <iso>] [--calendar <id|name>]",
      "  calendar create --subject \"...\" --start <iso> --end <iso> [--show-as free] [--execute]",
      "  calendar update <id> [--subject \"...\"] [--start <iso>] [--execute]",
      "  calendar cancel <id> [--comment \"...\"] [--execute --confirm <id>]",
      "  calendar delete <id> [--permanent] [--execute --confirm <id>]",
      "  calendar availability --schedules <a,b> [--start] [--end] [--interval N]",
      "  calendar suggest --attendees <a,b> [--duration 60] [--start] [--end]",
    ].join("\n"),
  ),
  "user": sub(
    { get: cmd(userGet), search: cmd(userSearch) },
    [
      "Valid subcommands:",
      "  user get <upn>            profile, contact info and manager",
      "  user search <term>        directory lookup by name, mail or upn prefix",
    ].join("\n"),
  ),
  "raw": cmd(raw),
};

const TOP_LEVEL_HELP = `msgraph-axi — TOON wrapper for Microsoft Graph (mail + calendar)

usage: msgraph-axi <command> [args] [flags]

commands:
  auth status|login|logout        sign-in state for the m365 backend
  mail list|read|send|delete|search|draft|thread|attachment
  calendar list|agenda|create|update|cancel|delete|availability|suggest
  user get <upn>|search <term>    people, manager and directory lookup
  raw <graph-path>                any Graph endpoint via m365 request

global flags: ${GLOBAL_FLAGS}

Run \`msgraph-axi <command> --help\` for command details.`;

const COMMAND_HELP: Record<string, string> = {
  auth: [
    "auth — m365 sign-in state (run this first)",
    "",
    "  msgraph-axi auth status            signed-in account and tenant",
    "  msgraph-axi auth login             interactive device-code sign-in",
    "  msgraph-axi auth login --auth-type certificate   service principal",
    "  msgraph-axi auth logout            sign out (no-op when signed out)",
    "",
    `flags: --auth-type <type>  ${GLOBAL_FLAGS}`,
  ].join("\n"),
  mail: [
    "mail — Outlook messages",
    "",
    "  msgraph-axi mail list [--folder inbox] [--start <iso>] [--end <iso>]",
    "      default fields: id, from, subject, received; --fields a,b to extend",
    "  msgraph-axi mail read <id> [--full]",
    "      default: snippet preview; --full adds the body content",
    "  msgraph-axi mail send --to a@x.com,b@y.com --subject \"...\" --body \"...\"",
    "      [--cc a] [--bcc b] [--body-type text|HTML] [--importance low|normal|high]",
    "      [--attach path,...] [--mailbox <upn>] [--execute]",
    "      without --execute prints a preview only (dry run)",
    "  msgraph-axi mail send --draft <id> [--execute]   send a saved draft",
    "  msgraph-axi mail delete <id> [--execute --confirm <id>]",
    "  msgraph-axi mail search --search \"<query>\" [--limit N]",
    "  msgraph-axi mail draft --to <a> --subject \"...\" --body \"...\" [--execute]",
    "  msgraph-axi mail thread <conversationId> [--full]",
    "  msgraph-axi mail attachment get <attachmentId> --message <messageId> [--out <path>]",
    "",
    `flags: ${GLOBAL_FLAGS}`,
  ].join("\n"),
  calendar: [
    "calendar — Outlook calendar and events",
    "",
    "  msgraph-axi calendar list",
    "  msgraph-axi calendar agenda [--start <iso>] [--end <iso>]",
    "      [--calendar <id|name>] [--timezone <tz>]   default: today..+7d",
    "  msgraph-axi calendar create --subject \"...\" --start <iso> --end <iso>",
    "      [--body \"...\"] [--location \"...\"] [--attendees a@x.com]",
    "      [--show-as free|tentative|busy|oof|workingElsewhere] [--execute]",
    "  msgraph-axi calendar update <id> [--subject \"...\"] [--start <iso>] [--show-as <state>] [--execute]",
    "  msgraph-axi calendar cancel <id> [--comment \"...\"] [--execute --confirm <id>]",
    "  msgraph-axi calendar delete <id> [--permanent] [--execute --confirm <id>]",
    "  msgraph-axi calendar availability --schedules a@x.com,b@y.com",
    "      [--start <iso>] [--end <iso>] [--interval 30] [--timezone <tz>]",
    "      availabilityView codes: 0=free 1=tentative 2=busy 3=oof 4=workingElsewhere",
    "  msgraph-axi calendar suggest --attendees a@x.com,b@y.com",
    "      [--duration 60] [--start <iso>] [--end <iso>] [--candidates 5] [--timezone <tz>]",
    "",
    "  --timezone expects an IANA name (Europe/Amsterdam) and defaults to your mailbox",
    "  time zone (this machine's zone as fallback);",
    "  --start/--end take a date, a wall clock in that zone, or an offset.",
    "  suggest (findMeetingTimes) can miss shared/delegated calendars: confirm a",
    "  slot with availability (getSchedule) before booking.",
    "",
    `flags: ${GLOBAL_FLAGS}`,
  ].join("\n"),
  user: [
    "user — people and organizational lookup via Graph",
    "",
    "  msgraph-axi user get <upn>",
    "      name, job title, department, location, contact info and manager",
    "  msgraph-axi user search <term> [--limit N] [--full]",
    "      directory lookup by display name, first/last name, mail or upn prefix",
    "      e.g. msgraph-axi user search alex",
    "",
    `flags: ${GLOBAL_FLAGS}`,
  ].join("\n"),
  raw: [
    "raw — any Microsoft Graph endpoint (escape hatch via m365 request)",
    "",
    "  msgraph-axi raw <path> [--query 'a=b&c=d']",
    "      e.g. msgraph-axi raw me/mailFolders/inbox?$select=unreadItemCount",
    "  msgraph-axi raw <path> --method post --body '{...}' [--content-type json]",
    "      write methods (post/put/patch/delete) require --execute",
    "  msgraph-axi raw <path> --method post --body @payload.json",
    "",
    `flags: --method <m> --body <json|@file> --content-type <ct> --prefer <h> --query <qs> ${GLOBAL_FLAGS}`,
  ].join("\n"),
};

export async function main(): Promise<void> {
  await runAxiCli<AppContext>({
    description: "Agent-ergonomic TOON wrapper for Microsoft Graph — mail and calendar on top of the CLI for Microsoft 365",
    version: VERSION,
    topLevelHelp: TOP_LEVEL_HELP,
    commands,
    home: cmd(home),
    getCommandHelp: (command) => COMMAND_HELP[command] ?? null,
    resolveContext: () => ({ m365: new PnpCliBackend() }),
  });
}