import { describe, expect, it } from "vitest";
import { AxiError } from "axi-sdk-js";
import { cleanupContext, expectAxiError, makeContext } from "./helpers.js";

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
});