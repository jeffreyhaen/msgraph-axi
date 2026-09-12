import { mkdtempSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { expect } from "vitest";
import { AxiError } from "axi-sdk-js";
import { PnpCliBackend } from "../src/backend.js";
import type { MailContext } from "../src/commands/mail.js";
import type { CalendarContext } from "../src/commands/calendar.js";
import type { AuthContext } from "../src/commands/auth.js";

export const FIXTURE = join(import.meta.dirname, "fixtures", "m365-cli.js");

export interface TestBackend extends MailContext, CalendarContext, AuthContext {
  m365: PnpCliBackend;
  logFile: string;
}

/** Backend that runs the fake m365 fixture and records every argv line. */
export function makeContext(
  env: Record<string, string> = {},
): TestBackend {
  const dir = mkdtempSync(join(tmpdir(), "msgraph-axi-test-"));
  const logFile = join(dir, "m365.log");
  return {
    m365: new PnpCliBackend(
      process.execPath,
      [FIXTURE],
      { FAKE_M365_LOG: logFile, ...env },
    ),
    user: undefined,
    logFile,
  };
}

/** The argv lines the fake m365 received, newest last. */
export function m365Calls(context: TestBackend): string[][] {
  try {
    const text = readFileSync(context.logFile, "utf8");
    return text
      .split("\n")
      .filter(Boolean)
      .map((line) => JSON.parse(line) as string[]);
  } catch {
    return [];
  }
}

/** The most recent m365 argv (skips the implicit status call). */
export function lastM365Call(context: TestBackend): string[] {
  const calls = m365Calls(context);
  return calls[calls.length - 1] ?? [];
}

export function cleanupContext(context: TestBackend): void {
  rmSync(join(context.logFile, ".."), { recursive: true, force: true });
}

/** Assert a rejected promise is an AxiError with code (+ optional message part). */
export async function expectAxiError(
  promise: Promise<unknown>,
  code: string,
  messagePart?: string,
): Promise<void> {
  try {
    await promise;
  } catch (error) {
    expect(error).toBeInstanceOf(AxiError);
    expect((error as AxiError).code).toBe(code);
    if (messagePart !== undefined) {
      expect((error as AxiError).message).toContain(messagePart);
    }
    return;
  }
  throw new Error(`Expected AxiError ${code} to be thrown`);
}