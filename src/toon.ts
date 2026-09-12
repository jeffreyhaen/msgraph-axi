const CELL_LIMIT = 200;

export interface ProjectionOptions {
  full: boolean;
}

/**
 * Truncate a single cell to CELL_LIMIT characters unless `--full` was given.
 */
export function cell(value: unknown, full: boolean): unknown {
  if (typeof value !== "string" || full) {
    return value;
  }
  if (value.length <= CELL_LIMIT) {
    return value;
  }
  return `${value.slice(0, CELL_LIMIT - 1)}\u2026`;
}

/**
 * Pick only the requested fields from an object, in the requested order,
 * applying per-cell truncation.
 */
export function project(
  item: Record<string, unknown>,
  fields: string[],
  full: boolean,
): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  for (const field of fields) {
    if (field in item) {
      out[field] = cell(item[field], full);
    }
  }
  return out;
}

export interface EmailAddress {
  emailAddress?: { name?: string; address?: string };
}

/** Compact `Name <address>` (or whichever part exists). */
export function formatFrom(value: EmailAddress | undefined): string {
  const inner = value?.emailAddress;
  if (!inner) {
    return "";
  }
  if (inner.name && inner.address) {
    return `${inner.name} <${inner.address}>`;
  }
  return inner.name ?? inner.address ?? "";
}

export function formatRecipients(
  values: EmailAddress[] | undefined,
): string {
  if (!values || values.length === 0) {
    return "";
  }
  return values.map(formatFrom).filter(Boolean).join("; ");
}

export interface GraphDateTime {
  dateTime?: string;
  timeZone?: string;
}

/** `2026-03-15T12:00:00 CET` — dateTime plus timeZone so the cell stays unambiguous. */
export function formatDateTime(value: GraphDateTime | undefined): string {
  if (!value?.dateTime) {
    return "";
  }
  if (value.timeZone) {
    return `${value.dateTime} ${value.timeZone}`;
  }
  return value.dateTime;
}

/** ISO timestamp (UTC, Z) for a local wall-clock midnight `daysFromNow` days out. */
export function localDayIso(daysFromNow: number, hours = 0): string {
  const now = new Date();
  const day = new Date(
    now.getFullYear(),
    now.getMonth(),
    now.getDate() + daysFromNow,
    hours,
    0,
    0,
    0,
  );
  return day.toISOString();
}

export function parseFields(
  fieldsFlag: string | undefined,
  defaults: string[],
): string[] {
  if (fieldsFlag === undefined || fieldsFlag.trim() === "") {
    return defaults;
  }
  return fieldsFlag
    .split(",")
    .map((f) => f.trim())
    .filter(Boolean);
}