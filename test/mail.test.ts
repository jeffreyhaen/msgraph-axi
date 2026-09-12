import { describe, expect, it } from "vitest";
import { AxiError } from "axi-sdk-js";
import { mailDelete, mailList, mailRead, mailSend } from "../src/commands/mail.js";
import { cleanupContext, expectAxiError, lastM365Call, makeContext, m365Calls } from "./helpers.js";

describe("mail list", () => {
  it("defaults to 20 rows with compact fields and truncation hints", async () => {
    const ctx = makeContext();
    try {
      const out = await mailList([], ctx);
      expect(out.count).toBe(20);
      expect(out.truncated).toBe(true);
      expect((out.help as string[])[0]).toContain("--limit");
      const rows = out.mail as Array<Record<string, unknown>>;
      expect(rows[0]).toMatchObject({
        id: "msg-1",
        from: "Alice <alice@contoso.com>",
        subject: "Subject 1",
        received: expect.stringMatching(/^2026-03-/),
      });
      expect(Object.keys(rows[0]).sort()).toEqual(["from", "id", "received", "subject"]);
    } finally {
      cleanupContext(ctx);
    }
  });

  it("honors --limit and --fields", async () => {
    const ctx = makeContext();
    try {
      const out = await mailList(["--limit", "5", "--fields", "isRead,hasAttachments"], ctx);
      expect(out.count).toBe(5);
      const rows = out.mail as Array<Record<string, unknown>>;
      expect(rows[2].isRead).toBe(true);
      expect(rows[2].hasAttachments).toBe(false);
    } finally {
      cleanupContext(ctx);
    }
  });

  it("passes folder, time window and user to m365", async () => {
    const ctx = makeContext();
    try {
      await mailList(
        [
          "--folder",
          "inbox",
          "--start",
          "2026-03-01T00:00:00Z",
          "--end",
          "2026-03-08T00:00:00Z",
          "--user",
          "bob@contoso.com",
        ],
        ctx,
      );
      const calls = m365Calls(ctx);
      const call = calls[calls.length - 1];
      expect(call.slice(0, 3)).toEqual(["outlook", "message", "list"]);
      expect(call).toEqual(
        expect.arrayContaining([
          "--folderName",
          "inbox",
          "--startTime",
          "2026-03-01T00:00:00Z",
          "--endTime",
          "2026-03-08T00:00:00Z",
          "--userName",
          "bob@contoso.com",
        ]),
      );
    } finally {
      cleanupContext(ctx);
    }
  });

  it("rejects unknown flags", async () => {
    const ctx = makeContext();
    try {
      await expectAxiError(mailList(["--bogus", "1"], ctx), "VALIDATION_ERROR", "--bogus");
    } finally {
      cleanupContext(ctx);
    }
  });
});

describe("mail read", () => {
  it("returns a compact snippet without the body", async () => {
    const ctx = makeContext();
    try {
      const out = await mailRead(["msg-1"], ctx);
      expect(out.id).toBe("msg-1");
      expect(out.to).toBe("Bob <bob@contoso.com>");
      expect(out.snippet).toContain("Preview of message 1");
      expect(out.body).toBeUndefined();
    } finally {
      cleanupContext(ctx);
    }
  });

  it("adds the body with --full", async () => {
    const ctx = makeContext();
    try {
      const out = await mailRead(["msg-1", "--full"], ctx);
      expect(out.body).toBe("Full body of message 1");
      expect(out.bodyContentType).toBe("text");
    } finally {
      cleanupContext(ctx);
    }
  });

  it("truncates long cells by default", async () => {
    const ctx = makeContext();
    try {
      const out = await mailRead(["msg-2"], ctx);
      expect((out.snippet as string).endsWith("\u2026")).toBe(true);
      const full = await mailRead(["msg-2", "--full"], ctx);
      expect((full.snippet as string).endsWith("\u2026")).toBe(false);
    } finally {
      cleanupContext(ctx);
    }
  });

  it("requires exactly one id", async () => {
    const ctx = makeContext();
    try {
      await expectAxiError(mailRead([], ctx), "VALIDATION_ERROR");
    } finally {
      cleanupContext(ctx);
    }
  });
});

describe("mail send", () => {
  it("dry-runs without --execute", async () => {
    const ctx = makeContext();
    try {
      const out = await mailSend(
        ["--to", "a@x.com,b@y.com", "--subject", "Hi", "--body", "Hello world"],
        ctx,
      );
      expect(out.execute).toBe(false);
      expect((out.preview as Record<string, unknown>).to).toBe("a@x.com,b@y.com");
      expect(m365Calls(ctx).length).toBe(0);
    } finally {
      cleanupContext(ctx);
    }
  });

  it("sends with --execute", async () => {
    const ctx = makeContext();
    try {
      const out = await mailSend(
        [
          "--to",
          "a@x.com",
          "--subject",
          "Hi",
          "--body",
          "Hello",
          "--importance",
          "high",
          "--execute",
        ],
        ctx,
      );
      expect(out).toMatchObject({ sent: true, subject: "Hi" });
      const call = lastM365Call(ctx);
      expect(call.slice(0, 3)).toEqual(["outlook", "mail", "send"]);
      expect(call).toEqual(
        expect.arrayContaining([
          "--to",
          "a@x.com",
          "--subject",
          "Hi",
          "--bodyContents",
          "Hello",
          "--importance",
          "high",
        ]),
      );
    } finally {
      cleanupContext(ctx);
    }
  });

  it("requires --to, --subject and --body", async () => {
    const ctx = makeContext();
    try {
      await expectAxiError(mailSend(["--to", "a@x.com"], ctx), "VALIDATION_ERROR");
    } finally {
      cleanupContext(ctx);
    }
  });
});

describe("mail delete", () => {
  it("requires --execute and an exact --confirm", async () => {
    const ctx = makeContext();
    try {
      const blocked = await mailDelete(["msg-1"], ctx);
      expect(blocked.execute).toBe(false);
      expect((blocked.help as string[])[0]).toContain("--confirm msg-1");
      expect((blocked.help as string[])[0]).toContain("--execute");

      const mismatched = await mailDelete(["msg-1", "--execute", "--confirm", "msg-2"], ctx);
      expect(mismatched.execute).toBe(false);
      expect(m365Calls(ctx).length).toBe(0);
    } finally {
      cleanupContext(ctx);
    }
  });

  it("deletes with --execute --confirm <id> and forwards --force", async () => {
    const ctx = makeContext();
    try {
      const out = await mailDelete(["msg-1", "--execute", "--confirm", "msg-1"], ctx);
      expect(out).toMatchObject({ deleted: true, id: "msg-1" });
      const call = lastM365Call(ctx);
      expect(call).toEqual(
        expect.arrayContaining([
          "outlook",
          "message",
          "remove",
          "--id",
          "msg-1",
          "--force",
          "--userName",
          "alice@contoso.com",
        ]),
      );
    } finally {
      cleanupContext(ctx);
    }
  });
});