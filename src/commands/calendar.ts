import { AxiError } from "axi-sdk-js";
import { PnpCliBackend, resolveBody } from "../backend.js";
import {
  parseFlags,
  flagString,
  flagNumber,
  flagBool,
  type FlagDef,
} from "../flags.js";
import {
  cell,
  formatDateTime,
  parseFields,
} from "../toon.js";
import {
  addDays,
  hasExplicitOffset,
  isBareDateTime,
  isDateOnly,
  localToUtcIso,
  wallClock,
} from "../time.js";

export interface CalendarContext {
  m365: PnpCliBackend;
  user?: string;
}

interface GraphEvent {
  id?: string;
  subject?: string;
  start?: { dateTime?: string; timeZone?: string };
  end?: { dateTime?: string; timeZone?: string };
  location?: { displayName?: string };
  [key: string]: unknown;
}

interface ScheduleAvailability {
  scheduleId?: string;
  availabilityView?: string;
  scheduleItems?: Array<{
    status?: string;
    start?: { dateTime?: string };
    end?: { dateTime?: string };
  }>;
}

const LIST_FLAGS: Record<string, FlagDef> = {
  user: { type: "string" },
  limit: { type: "number" },
  fields: { type: "string" },
  full: { type: "boolean" },
};

const AGENDA_FLAGS: Record<string, FlagDef> = {
  user: { type: "string" },
  start: { type: "string" },
  end: { type: "string" },
  calendar: { type: "string" },
  timezone: { type: "string" },
  limit: { type: "number" },
  fields: { type: "string" },
  full: { type: "boolean" },
};

const WRITE_FLAGS: Record<string, FlagDef> = {
  user: { type: "string" },
  subject: { type: "string" },
  start: { type: "string" },
  end: { type: "string" },
  body: { type: "string" },
  location: { type: "string" },
  attendees: { type: "string" },
  timezone: { type: "string" },
  calendar: { type: "string" },
  importance: { type: "string" },
  "show-as": { type: "string", aliases: ["showAs"] },
  execute: { type: "boolean" },
};

const DESTRUCTIVE_FLAGS: Record<string, FlagDef> = {
  user: { type: "string" },
  comment: { type: "string" },
  permanent: { type: "boolean" },
  execute: { type: "boolean" },
  confirm: { type: "string" },
};

const AVAIL_FLAGS: Record<string, FlagDef> = {
  user: { type: "string" },
  schedules: { type: "string" },
  start: { type: "string" },
  end: { type: "string" },
  interval: { type: "number" },
  timezone: { type: "string" },
  full: { type: "boolean" },
};

const SUGGEST_FLAGS: Record<string, FlagDef> = {
  user: { type: "string" },
  attendees: { type: "string" },
  duration: { type: "number" },
  start: { type: "string" },
  end: { type: "string" },
  candidates: { type: "number" },
  timezone: { type: "string" },
  full: { type: "boolean" },
};

const EVENT_DEFAULTS = ["id", "subject", "start", "end"];

export async function calendarList(
  args: string[],
  context: CalendarContext,
): Promise<Record<string, unknown>> {
  const parsed = parseFlags(args, LIST_FLAGS);
  const full = flagBool(parsed, "full", LIST_FLAGS);
  const limit = Math.floor(flagNumber(parsed, "limit", LIST_FLAGS, 20));
  const userArgs = await context.m365.userArgs(
    flagString(parsed, "user", LIST_FLAGS),
  );
  const items = await context.m365.runJsonArray<Record<string, unknown>>([
    "outlook",
    "calendar",
    "list",
    ...userArgs,
  ]);
  const shown = items.slice(0, limit);
  const fields = parseFields(flagString(parsed, "fields", LIST_FLAGS), [
    "id",
    "name",
  ]);
  const rows = shown.map((item) =>
    Object.fromEntries(
      fields.map((f) => [f, cell(item[f] ?? "", full)]),
    ),
  );
  const out: Record<string, unknown> = { calendars: rows, count: rows.length };
  if (shown.length < items.length) {
    out.truncated = true;
    out.help = [`Use --limit ${items.length} to see all ${items.length}`];
  }
  return out;
}

export async function calendarAgenda(
  args: string[],
  context: CalendarContext,
): Promise<Record<string, unknown>> {
  const parsed = parseFlags(args, AGENDA_FLAGS);
  const full = flagBool(parsed, "full", AGENDA_FLAGS);
  const limit = Math.floor(flagNumber(parsed, "limit", AGENDA_FLAGS, 20));
  const userFlag = flagString(parsed, "user", AGENDA_FLAGS);
  const timezone = await context.m365.timeZone(
    userFlag,
    flagString(parsed, "timezone", AGENDA_FLAGS),
  );
  const today = wallClock(new Date(), timezone).slice(0, 10);
  const start = agendaBound(
    flagString(parsed, "start", AGENDA_FLAGS),
    "start",
    timezone,
    localToUtcIso(`${today}T00:00:00`, timezone),
  );
  const end = agendaBound(
    flagString(parsed, "end", AGENDA_FLAGS),
    "end",
    timezone,
    localToUtcIso(`${addDays(today, 7)}T00:00:00`, timezone),
  );

  const m365Args = ["outlook", "event", "list"];
  const calendar = flagString(parsed, "calendar", AGENDA_FLAGS);
  if (calendar !== undefined) {
    m365Args.push(
      GUID_RE.test(calendar) ? "--calendarId" : "--calendarName",
      calendar,
    );
  }
  // The backend is strict about `--startDateTime`, so always hand it a zone.
  m365Args.push("--timeZone", timezone);
  m365Args.push(
    "--startDateTime",
    start,
    "--endDateTime",
    end,
    ...(await context.m365.userArgs(userFlag)),
  );

  const items = await context.m365.runJsonArray<GraphEvent>(m365Args);
  const sorted = [...items].sort((a, b) =>
    (a.start?.dateTime ?? "").localeCompare(b.start?.dateTime ?? ""),
  );
  const shown = sorted.slice(0, limit);
  const fields = parseFields(flagString(parsed, "fields", AGENDA_FLAGS), EVENT_DEFAULTS);
  const rows = shown.map((event) => {
    const row: Record<string, unknown> = {
      id: event.id ?? "",
      subject: cell(event.subject ?? "", full),
      start: formatDateTime(event.start),
      end: formatDateTime(event.end),
    };
    return { ...row, ...pickEventFields(event, fields, full) };
  });
  const out: Record<string, unknown> = {
    events: rows,
    count: rows.length,
    window: `${start}..${end}`,
    timezone,
  };
  if (shown.length < items.length) {
    out.truncated = true;
    out.help = [`Use --limit ${items.length} to see all ${items.length}`];
  }
  return out;
}

export async function calendarCreate(
  args: string[],
  context: CalendarContext,
): Promise<Record<string, unknown>> {
  const parsed = parseFlags(args, WRITE_FLAGS);
  const subject = flagString(parsed, "subject", WRITE_FLAGS);
  const start = flagString(parsed, "start", WRITE_FLAGS);
  const end = flagString(parsed, "end", WRITE_FLAGS);
  if (subject === undefined || start === undefined || end === undefined) {
    throw new AxiError(
      "calendar create requires --subject, --start and --end",
      "VALIDATION_ERROR",
      [
        'Example: msgraph-axi calendar create --subject "Review" --start 2026-03-15T12:00:00 --end 2026-03-15T13:00:00',
      ],
    );
  }
  const timezone = await context.m365.timeZone(
    flagString(parsed, "user", WRITE_FLAGS),
    flagString(parsed, "timezone", WRITE_FLAGS),
  );
  const body = buildEventBody(parsed, WRITE_FLAGS, timezone, true);
  const execute = flagBool(parsed, "execute", WRITE_FLAGS);
  if (!execute) {
    return {
      preview: { ...body, subject },
      execute: false,
      help: ["Run with --execute to create the event"],
    };
  }
  const id = await postEvent(
    context,
    flagString(parsed, "user", WRITE_FLAGS),
    flagString(parsed, "calendar", WRITE_FLAGS),
    body,
  );
  return { created: true, id, subject };
}

export async function calendarUpdate(
  args: string[],
  context: CalendarContext,
): Promise<Record<string, unknown>> {
  const parsed = parseFlags(args, WRITE_FLAGS);
  const positionals = parsed.positionals;
  if (positionals.length !== 1) {
    throw new AxiError(
      "calendar update expects exactly one event id",
      "VALIDATION_ERROR",
      ["Run `msgraph-axi calendar update <id> [--subject ...] [--start ...] ...`"],
    );
  }
  const id = positionals[0];
  const timezone = await context.m365.timeZone(
    flagString(parsed, "user", WRITE_FLAGS),
    flagString(parsed, "timezone", WRITE_FLAGS),
  );
  const body = buildEventBody(parsed, WRITE_FLAGS, timezone, false);
  if (Object.keys(body).length === 0) {
    throw new AxiError(
      "calendar update needs at least one field to change",
      "VALIDATION_ERROR",
      ["Example: msgraph-axi calendar update <id> --subject \"New title\""],
    );
  }
  const execute = flagBool(parsed, "execute", WRITE_FLAGS);
  if (!execute) {
    return {
      preview: { id, ...body },
      execute: false,
      help: ["Run with --execute to apply the change"],
    };
  }
  const userArgs = await context.m365.userArgs(
    flagString(parsed, "user", WRITE_FLAGS),
  );
  await context.m365.runJson([
    "request",
    "--method",
    "patch",
    "--url",
    eventUrl(userArgs, flagString(parsed, "calendar", WRITE_FLAGS), id),
    "--body",
    JSON.stringify(body),
    "--content-type",
    "application/json",
  ]);
  return { updated: true, id };
}

export async function calendarCancel(
  args: string[],
  context: CalendarContext,
): Promise<Record<string, unknown>> {
  const parsed = parseFlags(args, DESTRUCTIVE_FLAGS);
  const id = expectId(parsed, "calendar cancel");
  const execute = flagBool(parsed, "execute", DESTRUCTIVE_FLAGS);
  const confirm = flagString(parsed, "confirm", DESTRUCTIVE_FLAGS);
  if (!execute || confirm !== id) {
    return {
      destructive: true,
      id,
      execute: false,
      help: [`Run with --execute --confirm ${id} to cancel the event`],
    };
  }
  const m365Args = ["outlook", "event", "cancel", "--id", id, "--force"];
  const comment = flagString(parsed, "comment", DESTRUCTIVE_FLAGS);
  if (comment !== undefined) {
    m365Args.push("--comment", comment);
  }
  m365Args.push(
    ...(await context.m365.userArgs(flagString(parsed, "user", DESTRUCTIVE_FLAGS))),
  );
  await context.m365.run(m365Args);
  return { cancelled: true, id };
}

export async function calendarDelete(
  args: string[],
  context: CalendarContext,
): Promise<Record<string, unknown>> {
  const parsed = parseFlags(args, DESTRUCTIVE_FLAGS);
  const id = expectId(parsed, "calendar delete");
  const execute = flagBool(parsed, "execute", DESTRUCTIVE_FLAGS);
  const confirm = flagString(parsed, "confirm", DESTRUCTIVE_FLAGS);
  if (!execute || confirm !== id) {
    return {
      destructive: true,
      id,
      execute: false,
      help: [`Run with --execute --confirm ${id} to delete the event`],
    };
  }
  const m365Args = ["outlook", "event", "remove", "--id", id, "--force"];
  if (flagBool(parsed, "permanent", DESTRUCTIVE_FLAGS)) {
    m365Args.push("--permanent");
  }
  m365Args.push(
    ...(await context.m365.userArgs(flagString(parsed, "user", DESTRUCTIVE_FLAGS))),
  );
  await context.m365.run(m365Args);
  return { deleted: true, id, permanent: flagBool(parsed, "permanent", DESTRUCTIVE_FLAGS) };
}

export async function calendarAvailability(
  args: string[],
  context: CalendarContext,
): Promise<Record<string, unknown>> {
  const parsed = parseFlags(args, AVAIL_FLAGS);
  const full = flagBool(parsed, "full", AVAIL_FLAGS);
  const schedules = flagString(parsed, "schedules", AVAIL_FLAGS);
  if (schedules === undefined || schedules.trim() === "") {
    throw new AxiError(
      "calendar availability requires --schedules <upn1,upn2>",
      "VALIDATION_ERROR",
      ['Example: msgraph-axi calendar availability --schedules a@x.com,b@x.com'],
    );
  }
  const timezone = await context.m365.timeZone(
    flagString(parsed, "user", AVAIL_FLAGS),
    flagString(parsed, "timezone", AVAIL_FLAGS),
  );
  const interval = Math.floor(flagNumber(parsed, "interval", AVAIL_FLAGS, 30));
  const today = wallClock(new Date(), timezone).slice(0, 10);
  const body = {
    schedules: schedules.split(",").map((s) => s.trim()).filter(Boolean),
    startTime: {
      dateTime: graphBound(
        flagString(parsed, "start", AVAIL_FLAGS),
        "start",
        timezone,
        `${today}T08:00:00`,
      ),
      timeZone: timezone,
    },
    endTime: {
      dateTime: graphBound(
        flagString(parsed, "end", AVAIL_FLAGS),
        "end",
        timezone,
        `${today}T18:00:00`,
      ),
      timeZone: timezone,
    },
    availabilityViewInterval: interval,
  };
  const userArgs = await context.m365.userArgs(flagString(parsed, "user", AVAIL_FLAGS));
  const response = await context.m365.runJson<
    ScheduleAvailability[] | { value?: ScheduleAvailability[] } | { error: unknown }
  >([
    "request",
    "--method",
    "post",
    "--url",
    `${graphBase(userArgs)}/calendar/getSchedule`,
    "--body",
    JSON.stringify(body),
    "--content-type",
    "application/json",
  ]);
  const list = Array.isArray(response)
    ? response
    : Array.isArray((response as { value?: ScheduleAvailability[] }).value)
      ? (response as { value: ScheduleAvailability[] }).value
      : [];
  const rows = list.map((entry) => {
    const row: Record<string, unknown> = {
      schedule: entry.scheduleId ?? "",
      availabilityView: entry.availabilityView ?? "",
      busy: (entry.scheduleItems ?? []).filter(
        (item) => item.status && item.status !== "free",
      ).length,
    };
    if (full) {
      row.scheduleItems = (entry.scheduleItems ?? []).map((item) => ({
        status: item.status ?? "",
        start: item.start?.dateTime ?? "",
        end: item.end?.dateTime ?? "",
      }));
    }
    return row;
  });
  return {
    schedules: rows,
    count: rows.length,
    interval,
    timezone,
    legend: AVAILABILITY_LEGEND,
  };
}

interface MeetingSuggestion {
  meetingTimeSlot?: {
    start?: { dateTime?: string; timeZone?: string };
    end?: { dateTime?: string; timeZone?: string };
  };
  confidence?: number;
  organizerAvailability?: string;
  attendeeAvailability?: Array<{
    availability?: string;
    attendee?: { emailAddress?: { address?: string } };
  }>;
  suggestionReason?: string;
}

interface MeetingTimeResponse {
  meetingTimeSuggestions?: MeetingSuggestion[];
  emptySuggestionsReason?: string;
}

/**
 * Free-busy search across attendees via Graph `findMeetingTimes`.
 * Read-only; needs Calendars.Read (delegated) or Calendars.Read.Shared.
 */
export async function calendarSuggest(
  args: string[],
  context: CalendarContext,
): Promise<Record<string, unknown>> {
  const parsed = parseFlags(args, SUGGEST_FLAGS);
  const attendees = flagString(parsed, "attendees", SUGGEST_FLAGS);
  if (attendees === undefined || attendees.trim() === "") {
    throw new AxiError(
      "calendar suggest requires --attendees <upn1,upn2>",
      "VALIDATION_ERROR",
      ['Example: msgraph-axi calendar suggest --attendees a@x.com,b@x.com --duration 60'],
    );
  }
  const full = flagBool(parsed, "full", SUGGEST_FLAGS);
  const duration = Math.floor(flagNumber(parsed, "duration", SUGGEST_FLAGS, 60));
  const candidates = Math.floor(flagNumber(parsed, "candidates", SUGGEST_FLAGS, 5));
  const userFlag = flagString(parsed, "user", SUGGEST_FLAGS);
  const timezone = await context.m365.timeZone(
    userFlag,
    flagString(parsed, "timezone", SUGGEST_FLAGS),
  );
  const nowPlusHour = new Date(Date.now() + 60 * 60 * 1000);
  const start = graphBound(
    flagString(parsed, "start", SUGGEST_FLAGS),
    "start",
    timezone,
    `${wallClock(nowPlusHour, timezone).slice(0, 16)}:00`,
  );
  const end = graphBound(
    flagString(parsed, "end", SUGGEST_FLAGS),
    "end",
    timezone,
    wallClock(new Date(Date.now() + 7 * 24 * 60 * 60 * 1000), timezone),
  );
  const body = {
    attendees: attendees
      .split(",")
      .map((a) => a.trim())
      .filter(Boolean)
      .map((address) => ({ type: "required", emailAddress: { address } })),
    meetingDuration: `PT${duration}M`,
    maxCandidates: candidates,
    returnSuggestionReasons: true,
    timeConstraint: {
      timeslots: [
        {
          start: { dateTime: start, timeZone: timezone },
          end: { dateTime: end, timeZone: timezone },
        },
      ],
    },
  };
  const userArgs = await context.m365.userArgs(userFlag);
  const user = userArgs[1] ?? "me";
  const response = await context.m365.runJson<MeetingTimeResponse>([
    "request",
    "--method",
    "post",
    "--url",
    `@graph/users/${user}/findMeetingTimes`,
    "--body",
    JSON.stringify(body),
    "--content-type",
    "application/json",
  ]);
  const suggestions = response.meetingTimeSuggestions ?? [];
  if (suggestions.length === 0) {
    return {
      suggestions: [],
      count: 0,
      reason: response.emptySuggestionsReason ?? "no slots found",
    };
  }
  const rows = suggestions.map((suggestion) => {
    const slot = suggestion.meetingTimeSlot;
    const availability = suggestion.attendeeAvailability ?? [];
    const organizer = suggestion.organizerAvailability ?? "";
    // Graph reports the attendees only, so add the organizer explicitly and the
    // ratio reads as "everyone free" instead of "the one attendee is free".
    const attendeeFree = availability.filter(
      (a) => a.availability === "free" || a.availability === "unknown",
    ).length;
    const organizerFree = organizer === "free" || organizer === "unknown";
    const total = availability.length + (organizer === "" ? 0 : 1);
    const row: Record<string, unknown> = {
      start: slot?.start?.dateTime ?? "",
      end: slot?.end?.dateTime ?? "",
      confidence: `${Math.round(suggestion.confidence ?? 0)}%`,
      available: `${attendeeFree + (organizerFree ? 1 : 0)}/${total}`,
      organizer,
    };
    if (full) {
      row.reason = suggestion.suggestionReason ?? "";
      row.attendees = availability.map((a) => ({
        attendee: a.attendee?.emailAddress?.address ?? "",
        availability: a.availability ?? "",
      }));
    }
    return row;
  });
  return { suggestions: rows, count: rows.length, duration, timezone };
}

function buildEventBody(
  parsed: ReturnType<typeof parseFlags>,
  defs: Record<string, FlagDef>,
  timezone: string,
  includeAll: boolean,
): Record<string, unknown> {
  const timeZone = timezone;
  const body: Record<string, unknown> = {};
  const subject = flagString(parsed, "subject", defs);
  const start = flagString(parsed, "start", defs);
  const end = flagString(parsed, "end", defs);
  const location = flagString(parsed, "location", defs);
  const attendees = flagString(parsed, "attendees", defs);
  const importance = flagString(parsed, "importance", defs);
  const showAs = showAsValue(parsed, defs);
  const bodyFlag = flagString(parsed, "body", defs);

  if (subject !== undefined || includeAll) {
    body.subject = subject ?? "";
  }
  if (start !== undefined || includeAll) {
    body.start = { dateTime: graphBound(start, "start", timeZone), timeZone };
  }
  if (end !== undefined || includeAll) {
    body.end = { dateTime: graphBound(end, "end", timeZone), timeZone };
  }
  if (location !== undefined) {
    body.location = { displayName: location };
  }
  if (attendees !== undefined) {
    body.attendees = attendees
      .split(",")
      .map((a) => a.trim())
      .filter(Boolean)
      .map((address) => ({ emailAddress: { address }, type: "required" }));
  }
  if (importance !== undefined) {
    body.importance = importance;
  }
  if (showAs !== undefined) {
    body.showAs = showAs;
  }
  if (bodyFlag !== undefined) {
    body.body = { contentType: "text", content: resolveBody(bodyFlag) };
  }
  return body;
}

function pickEventFields(
  event: GraphEvent,
  fields: string[],
  full: boolean,
): Record<string, unknown> {
  // `fields` may extend the defaults with location or raw Graph fields.
  const out: Record<string, unknown> = {};
  for (const field of fields) {
    if (EVENT_DEFAULTS.includes(field)) {
      continue;
    }
    if (field === "location") {
      out.location = cell(event.location?.displayName ?? "", full);
      continue;
    }
    out[field] = cell(event[field] ?? "", full);
  }
  return out;
}

async function postEvent(
  context: CalendarContext,
  userFlag: string | undefined,
  calendarId: string | undefined,
  body: Record<string, unknown>,
): Promise<string | undefined> {
  const userArgs = await context.m365.userArgs(userFlag);
  const response = await context.m365.runJson<GraphEvent>([
    "request",
    "--method",
    "post",
    "--url",
    eventUrl(userArgs, calendarId),
    "--body",
    JSON.stringify(body),
    "--content-type",
    "application/json",
  ]);
  return response.id;
}

function eventUrl(
  userArgs: string[],
  calendarId: string | undefined,
  eventId?: string,
): string {
  const user = userArgs[1] ?? "me";
  const base =
    calendarId !== undefined
      ? `@graph/users/${user}/calendars/${calendarId}/events`
      : `@graph/users/${user}/events`;
  return eventId !== undefined ? `${base}/${eventId}` : base;
}

function expectId(
  parsed: ReturnType<typeof parseFlags>,
  command: string,
): string {
  if (parsed.positionals.length !== 1) {
    throw new AxiError(
      `${command} expects exactly one event id`,
      "VALIDATION_ERROR",
      [`Run \`msgraph-axi ${command} <id>\``],
    );
  }
  return parsed.positionals[0];
}

function graphBase(userArgs: string[]): string {
  const user = userArgs[1] ?? "me";
  return `@graph/users/${user}`;
}

const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/** Free/busy states Graph accepts for `showAs`. */
const SHOW_AS_VALUES = ["free", "tentative", "busy", "oof", "workingElsewhere"];

/** `getSchedule` availabilityView codes, so a reader needs no Graph docs. */
const AVAILABILITY_LEGEND = [
  "0=free",
  "1=tentative",
  "2=busy",
  "3=oof",
  "4=workingElsewhere",
];

const ISO_HINT =
  "Use ISO 8601: 2026-03-15T13:00:00 (wall clock in --timezone), 2026-03-15T13:00:00+02:00 or 2026-03-15";

function showAsValue(
  parsed: ReturnType<typeof parseFlags>,
  defs: Record<string, FlagDef>,
): string | undefined {
  const value = flagString(parsed, "show-as", defs) ?? flagString(parsed, "showAs", defs);
  if (value === undefined) {
    return undefined;
  }
  const normalized = value.trim();
  const match = SHOW_AS_VALUES.find(
    (candidate) => candidate.toLowerCase() === normalized.toLowerCase(),
  );
  if (match === undefined) {
    throw new AxiError(
      `--show-as expects one of ${SHOW_AS_VALUES.join(", ")} (got "${value}")`,
      "VALIDATION_ERROR",
      ["Use --show-as free to keep the slot visible as free/busy-neutral"],
    );
  }
  return match;
}

/**
 * Validate a `--start`/`--end` value. A bare wall clock is kept as-is for Graph
 * (which takes it together with `timeZone`) but converted to a `Z` instant for
 * the m365 backend, which rejects zone-less date-times.
 */
function bound(
  value: string,
  flag: string,
  timezone: string,
  toUtc: boolean,
): string {
  const trimmed = value.trim();
  if (isDateOnly(trimmed) || hasExplicitOffset(trimmed)) {
    return trimmed;
  }
  if (isBareDateTime(trimmed)) {
    return toUtc ? localToUtcIso(trimmed, timezone) : trimmed;
  }
  throw new AxiError(
    `--${flag} is not a valid ISO date-time: ${value}`,
    "VALIDATION_ERROR",
    [ISO_HINT],
  );
}

/** `--start`/`--end` for a Graph payload (wall clock plus `timeZone`). */
function graphBound(
  value: string | undefined,
  flag: string,
  timezone: string,
  fallback = "",
): string {
  return value === undefined ? fallback : bound(value, flag, timezone, false);
}

/** `--start`/`--end` for the m365 backend, whose date-times need an offset. */
function agendaBound(
  value: string | undefined,
  flag: string,
  timezone: string,
  fallback: string,
): string {
  return value === undefined ? fallback : bound(value, flag, timezone, true);
}