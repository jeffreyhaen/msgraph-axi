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
  localDayIso,
  parseFields,
} from "../toon.js";

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
  const start = flagString(parsed, "start", AGENDA_FLAGS) ?? localDayIso(0);
  const end = flagString(parsed, "end", AGENDA_FLAGS) ?? localDayIso(7);

  const m365Args = ["outlook", "event", "list"];
  const calendar = flagString(parsed, "calendar", AGENDA_FLAGS);
  if (calendar !== undefined) {
    m365Args.push(
      GUID_RE.test(calendar) ? "--calendarId" : "--calendarName",
      calendar,
    );
  }
  const timezone = flagString(parsed, "timezone", AGENDA_FLAGS);
  if (timezone !== undefined) {
    m365Args.push("--timeZone", timezone);
  }
  m365Args.push(
    "--startDateTime",
    start,
    "--endDateTime",
    end,
    ...(await context.m365.userArgs(flagString(parsed, "user", AGENDA_FLAGS))),
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
  const timezone = flagString(parsed, "timezone", WRITE_FLAGS);
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
  const timezone = flagString(parsed, "timezone", WRITE_FLAGS);
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
  const timezone = flagString(parsed, "timezone", AVAIL_FLAGS) ?? "UTC";
  const interval = Math.floor(flagNumber(parsed, "interval", AVAIL_FLAGS, 30));
  const body = {
    schedules: schedules.split(",").map((s) => s.trim()).filter(Boolean),
    startTime: {
      dateTime: flagString(parsed, "start", AVAIL_FLAGS) ?? localDayIso(0, 8),
      timeZone: timezone,
    },
    endTime: {
      dateTime: flagString(parsed, "end", AVAIL_FLAGS) ?? localDayIso(0, 18),
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
  return { schedules: rows, count: rows.length, interval };
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
  const timezone = flagString(parsed, "timezone", SUGGEST_FLAGS) ?? "UTC";
  const nowPlusHour = new Date(Date.now() + 60 * 60 * 1000);
  const start =
    flagString(parsed, "start", SUGGEST_FLAGS) ??
    `${nowPlusHour.toISOString().slice(0, 16)}:00`;
  const end =
    flagString(parsed, "end", SUGGEST_FLAGS) ??
    new Date(Date.now() + 7 * 24 * 60 * 60 * 1000).toISOString();
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
  const userArgs = await context.m365.userArgs(
    flagString(parsed, "user", SUGGEST_FLAGS),
  );
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
    const available = availability.filter(
      (a) => a.availability === "free" || a.availability === "unknown",
    ).length;
    const row: Record<string, unknown> = {
      start: slot?.start?.dateTime ?? "",
      end: slot?.end?.dateTime ?? "",
      confidence: `${Math.round((suggestion.confidence ?? 0) * 100)}%`,
      available: `${available}/${availability.length}`,
      organizer: suggestion.organizerAvailability ?? "",
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
  timezone: string | undefined,
  includeAll: boolean,
): Record<string, unknown> {
  const timeZone = timezone ?? "UTC";
  const body: Record<string, unknown> = {};
  const subject = flagString(parsed, "subject", defs);
  const start = flagString(parsed, "start", defs);
  const end = flagString(parsed, "end", defs);
  const location = flagString(parsed, "location", defs);
  const attendees = flagString(parsed, "attendees", defs);
  const importance = flagString(parsed, "importance", defs);
  const bodyFlag = flagString(parsed, "body", defs);

  if (subject !== undefined || includeAll) {
    body.subject = subject ?? "";
  }
  if (start !== undefined || includeAll) {
    body.start = { dateTime: start ?? "", timeZone };
  }
  if (end !== undefined || includeAll) {
    body.end = { dateTime: end ?? "", timeZone };
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