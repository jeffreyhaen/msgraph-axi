import { describe, expect, it } from "vitest";
import { home } from "../src/commands/home.js";
import { cleanupContext, makeContext } from "./helpers.js";

describe("home view", () => {
  it("shows the signed-in account with unread count and today's events", async () => {
    const ctx = makeContext();
    try {
      const out = await home([], ctx);
      expect(out.signedIn).toBe(true);
      expect(out.connectedAs).toBe("alice@contoso.com");
      expect(out.unreadMail).toBe(4);
      const today = out.today as Array<Record<string, unknown>>;
      expect(today.length).toBe(3);
      expect(today[0].id).toBeDefined();
    } finally {
      cleanupContext(ctx);
    }
  });

  it("degrades to an auth hint when signed out", async () => {
    const ctx = makeContext({ FAKE_M365_LOGGED_OUT: "1" });
    try {
      const out = await home([], ctx);
      expect(out.signedIn).toBe(false);
      expect((out.help as string[])[0]).toContain("auth login");
    } finally {
      cleanupContext(ctx);
    }
  });
});