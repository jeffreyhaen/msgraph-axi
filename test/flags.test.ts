import { describe, expect, it } from "vitest";
import { AxiError } from "axi-sdk-js";
import { parseFlags } from "../src/flags.js";

const DEFS = {
  limit: { type: "number" },
  fields: { type: "string" },
  full: { type: "boolean" },
  user: { type: "string", aliases: ["u"] },
} as const;

function expectValidationError(fn: () => unknown, messagePart?: string): AxiError {
  let caught: unknown;
  try {
    fn();
  } catch (error) {
    caught = error;
  }
  expect(caught).toBeInstanceOf(AxiError);
  const error = caught as AxiError;
  expect(error.code).toBe("VALIDATION_ERROR");
  if (messagePart !== undefined) {
    expect(error.message).toContain(messagePart);
  }
  return error;
}

describe("parseFlags", () => {
  it("splits positionals and flags", () => {
    const parsed = parseFlags(["msg-1", "--limit", "5", "--full"], DEFS);
    expect(parsed.positionals).toEqual(["msg-1"]);
    expect(parsed.flags).toEqual({ limit: 5, full: true });
  });

  it("supports = syntax and aliases, stored under the canonical name", () => {
    const parsed = parseFlags(["--limit=3", "--u", "a@x.com"], DEFS);
    expect(parsed.flags.limit).toBe(3);
    expect(parsed.flags.user).toBe("a@x.com");
    expect(parsed.flags.u).toBeUndefined();
  });

  it("rejects unknown flags with the valid list", () => {
    const error = expectValidationError(
      () => parseFlags(["--bogus", "1"], DEFS),
      "--bogus",
    );
    expect(error.suggestions.join(" ")).toContain("--limit");
  });

  it("rejects missing values", () => {
    expectValidationError(() => parseFlags(["--limit"], DEFS));
  });

  it("rejects non-numeric numbers", () => {
    expectValidationError(() => parseFlags(["--limit", "abc"], DEFS));
  });

  it("treats --flag false as boolean false", () => {
    const parsed = parseFlags(["--full=false"], DEFS);
    expect(parsed.flags.full).toBe(false);
  });
});