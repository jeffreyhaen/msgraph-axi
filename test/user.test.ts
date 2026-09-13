import { describe, expect, it } from "vitest";
import { userGet, userSearch } from "../src/commands/user.js";
import { cleanupContext, expectAxiError, lastM365Call, makeContext, m365Calls } from "./helpers.js";

describe("user get", () => {
  it("resolves the profile plus manager", async () => {
    const ctx = makeContext();
    try {
      const out = await userGet(["bob@contoso.com"], ctx);
      expect(out).toMatchObject({
        name: "Alice Wonder",
        jobTitle: "Engineer",
        department: "IT",
        location: "HQ-1",
        mail: "alice@contoso.com",
        upn: "alice@contoso.com",
        phone: "+1 555 0100",
        manager: "Carol Manager <carol@contoso.com>",
      });
      const calls = m365Calls(ctx);
      expect(calls.length).toBe(2); // user + manager (status is not called when targeting an explicit upn)
      const urls = calls.map((c) => c.join(" "));
      expect(urls.some((u) => u.includes(`/users/bob@contoso.com?$select=`))).toBe(true);
      expect(urls.some((u) => u.includes("/users/bob@contoso.com/manager?$select="))).toBe(true);
    } finally {
      cleanupContext(ctx);
    }
  });

  it("requires exactly one upn", async () => {
    const ctx = makeContext();
    try {
      await expectAxiError(userGet([], ctx), "VALIDATION_ERROR");
    } finally {
      cleanupContext(ctx);
    }
  });
});

describe("user search", () => {
  it("looks up people by name, mail or upn prefix", async () => {
    const ctx = makeContext();
    try {
      const out = await userSearch(["bob"], ctx);
      expect(out.count).toBe(2);
      const rows = out.users as Array<Record<string, unknown>>;
      expect(rows[0]).toMatchObject({
        name: "Alice Wonder",
        upn: "alice@contoso.com",
        mail: "alice@contoso.com",
        jobTitle: "Engineer",
      });
      const url = lastM365Call(ctx)[lastM365Call(ctx).indexOf("--url") + 1] ?? "";
      const decoded = decodeURIComponent(url);
      expect(decoded).toContain("@graph/users?$filter=");
      expect(decoded).toContain("startswith(displayName,'bob')");
      expect(decoded).toContain("startswith(surname,'bob')");
      expect(decoded).toContain("$top=15");
    } finally {
      cleanupContext(ctx);
    }
  });

  it("escapes quotes and honours --limit", async () => {
    const ctx = makeContext();
    try {
      await userSearch(["o'brien", "--limit", "5"], ctx);
      const decoded = decodeURIComponent(
        lastM365Call(ctx)[lastM365Call(ctx).indexOf("--url") + 1] ?? "",
      );
      expect(decoded).toContain("startswith(displayName,'o''brien')");
      expect(decoded).toContain("$top=5");
    } finally {
      cleanupContext(ctx);
    }
  });

  it("requires exactly one search term", async () => {
    const ctx = makeContext();
    try {
      await expectAxiError(userSearch([], ctx), "VALIDATION_ERROR");
      await expectAxiError(userSearch(["alex", "chen"], ctx), "VALIDATION_ERROR");
      expect(m365Calls(ctx).length).toBe(0);
    } finally {
      cleanupContext(ctx);
    }
  });
});