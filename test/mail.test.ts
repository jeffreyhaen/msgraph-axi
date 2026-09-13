import { describe, expect, it } from "vitest";
import { AxiError } from "axi-sdk-js";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
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

  it("passes folder, time window and user to Graph", async () => {
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
      expect(call.slice(0, 3)).toEqual(["request", "--method", "get"]);
      const url = call[call.indexOf("--url") + 1];
      expect(url).toContain("@graph/users/bob@contoso.com/mailFolders/inbox/messages");
      expect(url).toContain("$filter=");
      expect(decodeURIComponent(url)).toContain("receivedDateTime ge 2026-03-01T00:00:00Z");
      expect(decodeURIComponent(url)).toContain("receivedDateTime lt 2026-03-08T00:00:00Z");
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
  it("dry-runs without --execute and says it only saves a draft", async () => {
    const ctx = makeContext();
    try {
      const out = await mailSend(
        ["--to", "a@x.com,b@y.com", "--subject", "Hi", "--body", "Hello world"],
        ctx,
      );
      expect(out.execute).toBe(false);
      expect((out.preview as Record<string, unknown>).to).toBe("a@x.com,b@y.com");
      expect((out.preview as Record<string, unknown>).sends).toBe(false);
      expect((out.help as string[]).join(" ")).toContain("--send");
      expect(m365Calls(ctx).length).toBe(0);
    } finally {
      cleanupContext(ctx);
    }
  });

  it("saves a draft instead of delivering with --execute", async () => {
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
      expect(out).toMatchObject({ sent: false, draft: true, id: "draft-1" });
      const urls = m365Calls(ctx).map((call) => call.join(" "));
      expect(urls.some((url) => url.includes("/messages/draft-1/send"))).toBe(false);
      const call = lastM365Call(ctx);
      expect(call.slice(0, 3)).toEqual(["request", "--method", "post"]);
      const body = JSON.parse(call[call.indexOf("--body") + 1] ?? "{}") as Record<
        string,
        unknown
      >;
      expect(body).toMatchObject({
        subject: "Hi",
        body: { contentType: "text", content: "Hello" },
        importance: "high",
        isDraft: true,
      });
      expect((out.help as string[])[1]).toContain("mail send --draft draft-1 --execute");
    } finally {
      cleanupContext(ctx);
    }
  });

  it("delivers only with --send --execute", async () => {
    const ctx = makeContext();
    try {
      const out = await mailSend(
        ["--to", "a@x.com", "--subject", "Hi", "--body", "Hello", "--send", "--execute"],
        ctx,
      );
      expect(out).toMatchObject({ sent: true, draftId: "draft-1", to: "a@x.com" });
      const urls = m365Calls(ctx).map((call) => call.join(" "));
      expect(urls.some((url) => url.includes("/messages/draft-1/send"))).toBe(true);
    } finally {
      cleanupContext(ctx);
    }
  });

  it("keeps multi-line bodies intact", async () => {
    const ctx = makeContext();
    const body = ["Hi Alex,", "", "The repository is on GitHub.", "", "Regards,", "Sam"].join("\n");
    try {
      await mailSend(
        ["--to", "a@x.com", "--subject", "Hi", "--body", body, "--execute"],
        ctx,
      );
      const call = lastM365Call(ctx);
      const argument = call[call.indexOf("--body") + 1] ?? "";
      // The payload must stay JSON: a literal line break in an argv value is
      // silently truncated by cmd.exe / cross-spawn on Windows.
      expect(argument).not.toMatch(/[\r\n]/);
      const payload = JSON.parse(argument) as Record<string, unknown>;
      expect((payload.body as Record<string, unknown>).content).toBe(body);
    } finally {
      cleanupContext(ctx);
    }
  });

  it("attaches files to the draft", async () => {
    const ctx = makeContext();
    const dir = mkdtempSync(join(tmpdir(), "msgraph-axi-mail-"));
    const file = join(dir, "report.csv");
    writeFileSync(file, "a,b\n1,2\n", "utf8");
    try {
      await mailSend(
        [
          "--to",
          "a@x.com",
          "--subject",
          "Hi",
          "--body",
          "See the attachment",
          "--attach",
          file,
          "--execute",
        ],
        ctx,
      );
      const call = lastM365Call(ctx);
      const payload = JSON.parse(call[call.indexOf("--body") + 1] ?? "{}") as Record<
        string,
        unknown
      >;
      const attachments = payload.attachments as Array<Record<string, string>>;
      expect(attachments[0]).toMatchObject({
        name: "report.csv",
        contentType: "text/csv",
      });
      expect(Buffer.from(attachments[0].contentBytes, "base64").toString("utf8")).toBe(
        "a,b\n1,2\n",
      );
    } finally {
      cleanupContext(ctx);
      rmSync(dir, { recursive: true, force: true });
    }
  });

  it("reports a missing attachment before touching Graph", async () => {
    const ctx = makeContext();
    try {
      await expectAxiError(
        mailSend(
          ["--to", "a@x.com", "--subject", "Hi", "--body", "x", "--attach", "nope.pdf", "--execute"],
          ctx,
        ),
        "VALIDATION_ERROR",
        "Attachment not found",
      );
      expect(m365Calls(ctx).length).toBe(0);
    } finally {
      cleanupContext(ctx);
    }
  });

  it("delivers an existing draft with --draft", async () => {
    const ctx = makeContext();
    try {
      const out = await mailSend(["--draft", "draft-7", "--execute"], ctx);
      expect(out).toMatchObject({ sent: true, draftId: "draft-7" });
      expect(lastM365Call(ctx).join(" ")).toContain("/messages/draft-7/send");
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