const DATE_ONLY_RE = /^\d{4}-\d{2}-\d{2}$/;
const BARE_DATETIME_RE = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}(:\d{2})?(\.\d+)?$/;
const EXPLICIT_OFFSET_RE = /(?:Z|[+-]\d{2}:?\d{2})$/i;

/** `2026-03-15` — a date without a time. */
export function isDateOnly(value: string): boolean {
  return DATE_ONLY_RE.test(value.trim());
}

/** `2026-03-15T13:00[:00]` — a wall-clock time without an offset. */
export function isBareDateTime(value: string): boolean {
  return BARE_DATETIME_RE.test(value.trim());
}

/** Ends with `Z` or a `+hh:mm`/`-hhmm` offset, so the instant is unambiguous. */
export function hasExplicitOffset(value: string): boolean {
  const trimmed = value.trim();
  // Require a time part: `15-09-2026` must not read as an offset of -2026.
  return trimmed.includes("T") && EXPLICIT_OFFSET_RE.test(trimmed);
}

/** This machine's IANA time zone, or `UTC` when the runtime cannot report one. */
export function hostTimeZone(): string {
  try {
    const zone = Intl.DateTimeFormat().resolvedOptions().timeZone;
    return zone && zone.trim() !== "" ? zone : "UTC";
  } catch {
    return "UTC";
  }
}

interface WallParts {
  year: number;
  month: number;
  day: number;
  hour: number;
  minute: number;
  second: number;
}

function wallParts(instant: Date, timeZone: string): WallParts {
  const parts = new Intl.DateTimeFormat("en-US", {
    timeZone,
    hourCycle: "h23",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
  }).formatToParts(instant);
  const get = (type: string): number => {
    const value = parts.find((part) => part.type === type)?.value;
    return value === undefined ? 0 : Number(value);
  };
  return {
    year: get("year"),
    month: get("month"),
    day: get("day"),
    hour: get("hour"),
    minute: get("minute"),
    second: get("second"),
  };
}

/** Offset of `timeZone` from UTC, in minutes, at the given instant. */
function zoneOffsetMinutes(instant: Date, timeZone: string): number {
  const wall = wallParts(instant, timeZone);
  const asUtc = Date.UTC(
    wall.year,
    wall.month - 1,
    wall.day,
    wall.hour,
    wall.minute,
    wall.second,
  );
  const truncated = Math.floor(instant.getTime() / 1000) * 1000;
  return Math.round((asUtc - truncated) / 60000);
}

/**
 * Convert a wall-clock ISO string (`2026-03-15T13:00:00`, no offset) into a UTC
 * `...Z` instant, interpreting the wall clock in `timeZone`. Unparseable input
 * and unknown time zones are returned unchanged so callers can pass the value
 * through to a backend that accepts more formats.
 */
export function localToUtcIso(localIso: string, timeZone: string): string {
  const trimmed = localIso.trim();
  const naive = new Date(`${trimmed}Z`);
  if (Number.isNaN(naive.getTime())) {
    return trimmed;
  }
  try {
    let offset = zoneOffsetMinutes(naive, timeZone);
    let instant = new Date(naive.getTime() - offset * 60000);
    // A DST boundary can change the offset between the two guesses.
    const corrected = zoneOffsetMinutes(instant, timeZone);
    if (corrected !== offset) {
      offset = corrected;
      instant = new Date(naive.getTime() - offset * 60000);
    }
    return instant.toISOString();
  } catch {
    return trimmed;
  }
}

/** `2026-03-15T13:00:00` — the wall clock of `instant` in `timeZone`. */
export function wallClock(instant: Date, timeZone: string): string {
  try {
    const wall = wallParts(instant, timeZone);
    const pad = (n: number): string => String(n).padStart(2, "0");
    return `${wall.year}-${pad(wall.month)}-${pad(wall.day)}T${pad(wall.hour)}:${pad(wall.minute)}:${pad(wall.second)}`;
  } catch {
    return instant.toISOString().slice(0, 19);
  }
}

/** Add whole days to a `2026-03-15` date string. */
export function addDays(dateOnly: string, days: number): string {
  const [year, month, day] = dateOnly.split("-").map(Number);
  const shifted = new Date(Date.UTC(year, (month ?? 1) - 1, (day ?? 1) + days));
  return shifted.toISOString().slice(0, 10);
}
