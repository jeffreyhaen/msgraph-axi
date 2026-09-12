import { describe, expect, it } from "vitest";
import { userGet } from "../src/commands/user.js";
import { cleanupContext, expectAxiError, makeContext, m365Calls } from "./helpers.js";

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
      expect(calls.length).toBe(3); // status + user + manager
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