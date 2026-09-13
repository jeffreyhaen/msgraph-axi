import { describe, expect, it } from "vitest";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { raw } from "../src/commands/raw.js";
import { cleanupContext, expectAxiError, lastM365Call, makeContext, m365Calls } from "./helpers.js";

describe("raw", () => {
  it("gets a Graph path through m365 request", async () => {
    const ctx = makeContext();
    try {
      const out = (await raw(["me/mailFolders/inbox?$select=unreadItemCount"], ctx)) as Record<
        string,
        unknown
      >;
      expect(out).toMatchObject({ unreadItemCount: 4 });
      const call = lastM365Call(ctx);
      expect(call).toEqual(
        expect.arrayContaining([
          "request",
          "--method",
          "get",
          "--url",
          "@graph/me/mailFolders/inbox?$select=unreadItemCount",
        ]),
      );
    } finally {
      cleanupContext(ctx);
    }
  });

  it("appends --query to the URL", async () => {
    const ctx = makeContext();
    try {
      await raw(["me/events", "--query", "$top=5"], ctx);
      const call = lastM365Call(ctx);
      expect(call.join(" ")).toContain("@graph/me/events?$top=5");
    } finally {
      cleanupContext(ctx);
    }
  });

  it("gates write methods behind --execute", async () => {
    const ctx = makeContext();
    try {
      const blocked = (await raw(
        ["me/events", "--method", "post", "--body", "{}"],
        ctx,
      )) as Record<string, unknown>;
      expect(blocked.execute).toBe(false);
      expect(m365Calls(ctx).length).toBe(0);

      await raw(["me/events", "--method", "post", "--body", "{}", "--execute"], ctx);
      const call = lastM365Call(ctx);
      expect(call.slice(0, 3)).toEqual(["request", "--method", "post"]);
      expect(call).toEqual(expect.arrayContaining(["--content-type", "application/json"]));
    } finally {
      cleanupContext(ctx);
    }
  });

  it("rejects --body on get", async () => {
    const ctx = makeContext();
    try {
      await expectAxiError(raw(["me/events", "--body", "{}"], ctx), "VALIDATION_ERROR");
    } finally {
      cleanupContext(ctx);
    }
  });

  it("rejects unknown methods and missing paths", async () => {
    const ctx = makeContext();
    try {
      await expectAxiError(raw(["x", "--method", "teapot"], ctx), "VALIDATION_ERROR", "teapot");
      await expectAxiError(raw([], ctx), "VALIDATION_ERROR");
    } finally {
      cleanupContext(ctx);
    }
  });

  it("hands @file bodies to the backend instead of inlining them", async () => {
    const ctx = makeContext();
    const dir = mkdtempSync(join(tmpdir(), "msgraph-axi-raw-"));
    const file = join(dir, "payload.json");
    writeFileSync(file, '{\n  "subject": "Hi"\n}\n', "utf8");
    try {
      await raw(["me/messages", "--method", "post", "--body", `@${file}`, "--execute"], ctx);
      const call = lastM365Call(ctx);
      expect(call[call.indexOf("--body") + 1]).toBe(`@${file}`);
    } finally {
      cleanupContext(ctx);
      rmSync(dir, { recursive: true, force: true });
    }
  });
});