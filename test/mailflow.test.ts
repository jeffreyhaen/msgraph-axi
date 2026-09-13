import { describe, expect, it } from "vitest";
import { existsSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import {
  mailAttachmentGet,
  mailDraft,
  mailSearch,
  mailThread,
} from "../src/commands/mailflow.js";
import { mailSend } from "../src/commands/mail.js";
import {
  cleanupContext,
  expectAxiError,
  lastM365Call,
  makeContext,
  requestBody,
} from "./helpers.js";

describe("mail search", () => {
  it("requires --search", async () => {
    const ctx = makeContext();
    try {
      await expectAxiError(mailSearch([], ctx), "VALIDATION_ERROR");
    } finally {
      cleanupContext(ctx);
    }
  });

  it("maps Graph search results to compact rows", async () => {
    const ctx = makeContext();
    try {
      const out = await mailSearch(["--search", "invoice overdue"], ctx);
      const rows = out.mail as Array<Record<string, unknown>>;
      expect(rows.length).toBe(3);
      expect(rows[0]).toMatchObject({
        id: "msg-1",
        from: "Alice <alice@contoso.com>",
        subject: "Subject 1",
      });
      expect(out.query).toBe("invoice overdue");
      const call = lastM365Call(ctx);
      expect(call.join(" ")).toContain("/messages?$search=");
      expect(call.join(" ")).toContain("%22invoice");
    } finally {
      cleanupContext(ctx);
    }
  });
});

describe("mail draft", () => {
  it("dry-runs without --execute", async () => {
    const ctx = makeContext();
    try {
      const out = await mailDraft(
        ["--to", "a@x.com", "--subject", "Hi", "--body", "Hello"],
        ctx,
      );
      expect(out.execute).toBe(false);
      expect((out.preview as Record<string, unknown>).to).toBe("a@x.com");
    } finally {
      cleanupContext(ctx);
    }
  });

  it("creates an isDraft message with --execute", async () => {
    const ctx = makeContext();
    try {
      const out = await mailDraft(
        [
          "--to",
          "a@x.com,b@y.com",
          "--cc",
          "c@x.com",
          "--subject",
          "Hi",
          "--body",
          "Hello",
          "--execute",
        ],
        ctx,
      );
      expect(out).toMatchObject({ draft: true, id: "draft-1", subject: "Hi" });
      const call = lastM365Call(ctx);
      const body = requestBody(call);
      expect(body.isDraft).toBe(true);
      expect((body.toRecipients as Array<Record<string, unknown>>).length).toBe(2);
      expect((body.ccRecipients as Array<Record<string, unknown>>).length).toBe(1);
      expect(call.join(" ")).toContain("@graph/users/alice@contoso.com/messages");
    } finally {
      cleanupContext(ctx);
    }
  });

  it("requires --to, --subject and --body", async () => {
    const ctx = makeContext();
    try {
      await expectAxiError(mailDraft(["--to", "a@x.com"], ctx), "VALIDATION_ERROR");
    } finally {
      cleanupContext(ctx);
    }
  });
});

describe("mail send --draft", () => {
  it("dry-runs without --execute", async () => {
    const ctx = makeContext();
    try {
      const out = await mailSend(["--draft", "draft-1"], ctx);
      expect(out.execute).toBe(false);
      expect((out.preview as Record<string, unknown>).draftId).toBe("draft-1");
    } finally {
      cleanupContext(ctx);
    }
  });

  it("posts to /messages/{id}/send with --execute", async () => {
    const ctx = makeContext();
    try {
      const out = await mailSend(["--draft", "draft-1", "--execute"], ctx);
      expect(out).toMatchObject({ sent: true, draftId: "draft-1" });
      const call = lastM365Call(ctx);
      expect(call.join(" ")).toContain(
        "@graph/users/alice@contoso.com/messages/draft-1/send",
      );
    } finally {
      cleanupContext(ctx);
    }
  });

  it("rejects combining --draft with a new message", async () => {
    const ctx = makeContext();
    try {
      await expectAxiError(
        mailSend(["--draft", "draft-1", "--subject", "X"], ctx),
        "VALIDATION_ERROR",
      );
    } finally {
      cleanupContext(ctx);
    }
  });
});

describe("mail thread", () => {
  it("returns the conversation sorted oldest-first", async () => {
    const ctx = makeContext();
    try {
      const out = await mailThread(["conv-42"], ctx);
      expect(out.count).toBe(3);
      expect(out.conversationId).toBe("conv-42");
      const rows = out.thread as Array<Record<string, unknown>>;
      expect(rows[0].id).toBe("t1");
      expect(rows[2].id).toBe("t3");
      expect(rows[0].from).toBe("Alice <alice@contoso.com>");
      const call = lastM365Call(ctx);
      expect(call.join(" ")).toContain("$filter=conversationId");
    } finally {
      cleanupContext(ctx);
    }
  });

  it("requires exactly one conversation id", async () => {
    const ctx = makeContext();
    try {
      await expectAxiError(mailThread([], ctx), "VALIDATION_ERROR");
    } finally {
      cleanupContext(ctx);
    }
  });
});

describe("mail attachment get", () => {
  it("downloads the attachment to stdout-name by default", async () => {
    const ctx = makeContext();
    try {
      const out = await mailAttachmentGet(
        ["att-9", "--message", "msg-1"],
        ctx,
      );
      const saved = out.saved as string;
      expect(saved.endsWith("report.pdf")).toBe(true);
      expect(readFileSync(saved, "utf8")).toBe("pdf-bytes");
      expect(out).toMatchObject({ attachment: "report.pdf", size: 9 });
      rmSync(saved, { force: true });
    } finally {
      cleanupContext(ctx);
    }
  });

  it("honors --out and requires --message", async () => {
    const ctx = makeContext();
    const outPath = join(tmpdir(), "msgraph-axi-download-test.pdf");
    try {
      const out = await mailAttachmentGet(
        ["att-9", "--message", "msg-1", "--out", outPath],
        ctx,
      );
      expect(out.saved).toBe(outPath);
      expect(existsSync(outPath)).toBe(true);
      await expectAxiError(
        mailAttachmentGet(["att-9"], ctx),
        "VALIDATION_ERROR",
      );
    } finally {
      cleanupContext(ctx);
    }
  });
});