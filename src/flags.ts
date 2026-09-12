import { AxiError } from "axi-sdk-js";

export interface FlagDef {
  type: "string" | "boolean" | "number";
  aliases?: string[];
}

export interface ParsedArgs {
  positionals: string[];
  flags: Record<string, string | number | boolean>;
}

const isFlag = (arg: string): boolean => arg.startsWith("--");

export function parseFlags(
  args: string[],
  defs: Record<string, FlagDef>,
): ParsedArgs {
  const positionals: string[] = [];
  const flags: Record<string, string | number | boolean> = {};
  const valid = new Map<string, FlagDef>();
  for (const [name, def] of Object.entries(defs)) {
    valid.set(name, def);
    for (const alias of def.aliases ?? []) {
      valid.set(alias, def);
    }
  }

  const unknown: string[] = [];
  for (let i = 0; i < args.length; i++) {
    const arg = args[i];
    if (!isFlag(arg)) {
      positionals.push(arg);
      continue;
    }
    const eq = arg.indexOf("=");
    const rawName = eq >= 0 ? arg.slice(2, eq) : arg.slice(2);
    const inlineValue = eq >= 0 ? arg.slice(eq + 1) : undefined;
    const def = valid.get(rawName);
    if (!def) {
      unknown.push(arg);
      continue;
    }
    const canonical = canonicalOf(rawName);
    if (def.type === "boolean") {
      flags[canonical] =
        inlineValue !== undefined ? inlineValue !== "false" : true;
      continue;
    }
    let value: string | undefined = inlineValue;
    if (value === undefined) {
      const next = args[i + 1];
      if (next !== undefined && !isFlag(next)) {
        value = next;
        i++;
      }
    }
    if (value === undefined) {
      throw new AxiError(`Flag --${canonical} requires a value`, "VALIDATION_ERROR", [
        `Run with --${canonical} <value>`,
      ]);
    }
    flags[canonical] =
      def.type === "number" ? parseNumber(canonical, value) : value;
  }

  function canonicalOf(name: string): string {
    for (const n of Object.keys(defs)) {
      const d = valid.get(n);
      if (n === name || (d?.aliases ?? []).includes(name)) {
        return n;
      }
    }
    return name;
  }

  if (unknown.length > 0) {
    throw new AxiError(
      `Unknown flag${unknown.length > 1 ? "s" : ""}: ${unknown.join(", ")}`,
      "VALIDATION_ERROR",
      [`Valid flags: ${Object.keys(defs).map((n) => `--${n}`).join(" ")}`],
    );
  }

  return { positionals, flags };
}

function parseNumber(name: string, value: string): number {
  const n = Number(value);
  if (!Number.isFinite(n) || n < 0) {
    throw new AxiError(
      `Flag --${name} expects a non-negative number`,
      "VALIDATION_ERROR",
      [`Use --${name} <number>`],
    );
  }
  return n;
}

export function flagString(
  parsed: ParsedArgs,
  name: string,
  defs: Record<string, FlagDef>,
): string | undefined {
  const canonical = canonicalName(name, defs);
  const value = parsed.flags[canonical];
  return typeof value === "string" ? value : undefined;
}

export function flagNumber(
  parsed: ParsedArgs,
  name: string,
  defs: Record<string, FlagDef>,
  fallback: number,
): number {
  const canonical = canonicalName(name, defs);
  const value = parsed.flags[canonical];
  return typeof value === "number" ? value : fallback;
}

export function flagBool(
  parsed: ParsedArgs,
  name: string,
  defs: Record<string, FlagDef>,
): boolean {
  const canonical = canonicalName(name, defs);
  return parsed.flags[canonical] === true;
}

function canonicalName(name: string, defs: Record<string, FlagDef>): string {
  for (const [n, def] of Object.entries(defs)) {
    if (n === name || (def.aliases ?? []).includes(name)) {
      return n;
    }
  }
  return name;
}