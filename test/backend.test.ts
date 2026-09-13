import { describe, expect, it } from "vitest";
import { AxiError } from "axi-sdk-js";
import { cleanupContext, expectAxiError, makeContext } from "./helpers.js";

/** Run a command against a fixture that fails with the given message. */
async function failureFor(message: string): Promise<AxiError> {
  const ctx = makeContext({ FAKE_M365_ERROR: message });
  try {
    return await ctx.m365
      .run(["outlook", "event", "list"])
      .then(
        () => {
          throw new Error("expected m365 to fail");
        },
        (error: AxiError) => error,
      );
  } finally {
    cleanupContext(ctx);
  }
}

describe("PnpCliBackend", () => {
  it("parses m365 status into the signed-in account", async () => {
    const ctx = makeContext();
    try {
      const status = await ctx.m365.status();
      expect(status?.connectedAs).toBe("alice@contoso.com");
      expect(status?.authType).toBe("deviceCode");
    } finally {
      cleanupContext(ctx);
    }
  });

  it("treats a logged-out status as undefined, not an error", async () => {
    const ctx = makeContext({ FAKE_M365_LOGGED_OUT: "1" });
    try {
      expect(await ctx.m365.status()).toBeUndefined();
    } finally {
      cleanupContext(ctx);
    }
  });

  it("resolves the current user as --userName", async () => {
    const ctx = makeContext();
    try {
      expect(await ctx.m365.userArgs(undefined)).toEqual([
        "--userName",
        "alice@contoso.com",
      ]);
    } finally {
      cleanupContext(ctx);
    }
  });

  it("passes a GUID --user through as --userId", async () => {
    const ctx = makeContext();
    try {
      expect(await ctx.m365.userArgs("123e4567-e89b-12d3-a456-426614174000")).toEqual([
        "--userId",
        "123e4567-e89b-12d3-a456-426614174000",
      ]);
    } finally {
      cleanupContext(ctx);
    }
  });

  it("maps m365 runtime errors to M365_ERROR", async () => {
    const ctx = makeContext();
    try {
      await expect(ctx.m365.run(["outlook", "message", "get", "--id", "missing"])).rejects.toMatchObject({
        code: "M365_ERROR",
      });
    } finally {
      cleanupContext(ctx);
    }
  });

  it("maps sign-in failures to NOT_SIGNED_IN", async () => {
    const ctx = makeContext({ FAKE_M365_NOT_SIGNED_IN: "1" });
    try {
      await expect(ctx.m365.run(["outlook", "message", "list"])).rejects.toMatchObject({
        code: "NOT_SIGNED_IN",
      });
    } finally {
      cleanupContext(ctx);
    }
  });

  it("rejects non-array JSON from runJsonArray", async () => {
    const ctx = makeContext();
    try {
      await expectAxiError(
        ctx.m365.runJsonArray(["request", "--method", "get", "--url", "@graph/x"]),
        "M365_UNEXPECTED_OUTPUT",
      );
    } finally {
      cleanupContext(ctx);
    }
  });

  it("points at the input, not the connection, for a bad parameter", async () => {
    const error = await failureFor(
      "startDateTime: '2026-09-15T00:00:00' is not a valid ISO date-time.",
    );
    expect(error.code).toBe("M365_ERROR");
    const hints = error.suggestions.join(" ");
    expect(hints).toContain("input error");
    expect(hints).not.toContain("auth status");
  });

  it("points at missing permissions for a 403", async () => {
    const error = await failureFor("Request failed with status code 403");
    const hints = error.suggestions.join(" ");
    expect(hints).toMatch(/permission/i);
    expect(hints).not.toContain("auth status");
  });

  it("still suggests signing in for an expired session", async () => {
    const error = await failureFor("Request failed with status code 401");
    expect(error.suggestions.join(" ")).toContain("auth login");
  });
});