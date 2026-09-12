import { AxiError } from "axi-sdk-js";
import { PnpCliBackend } from "../backend.js";
import { parseFlags, flagString, flagNumber, flagBool, type FlagDef } from "../flags.js";
import {
  cell,
  formatFrom,
  formatRecipients,
  parseFields,
  project,
  type EmailAddress,
} from "../toon.js";
import { mailSendDraft } from "./mailflow.js";

export interface MailContext {
  m365: PnpCliBackend;
  user?: string;
}

export interface MailListItem {
  id?: string;
  subject?: string;
  from?: { emailAddress?: { name?: string; address?: string } };
  receivedDateTime?: string;
  isRead?: boolean;
  hasAttachments?: boolean;
  [key: string]: unknown;
}

export interface MessageDetail extends MailListItem {
  toRecipients?: unknown[];
  bodyPreview?: string;
  body?: { contentType?: string; content?: string };
  attachments?: unknown[];
}

const LIST_FLAGS: Record<string, FlagDef> = {
  folder: { type: "string" },
  start: { type: "string" },
  end: { type: "string" },
  user: { type: "string" },
  limit: { type: "number" },
  fields: { type: "string" },
  full: { type: "boolean" },
};

const READ_FLAGS: Record<string, FlagDef> = {
  user: { type: "string" },
  full: { type: "boolean" },
};

const SEND_FLAGS: Record<string, FlagDef> = {
  to: { type: "string" },
  cc: { type: "string" },
  bcc: { type: "string" },
  subject: { type: "string", aliases: ["s"] },
  body: { type: "string" },
  "body-type": { type: "string" },
  importance: { type: "string" },
  attach: { type: "string" },
  mailbox: { type: "string", aliases: ["m"] },
  sender: { type: "string" },
  draft: { type: "string" },
  execute: { type: "boolean" },
};

const DELETE_FLAGS: Record<string, FlagDef> = {
  user: { type: "string" },
  execute: { type: "boolean" },
  confirm: { type: "string" },
};

const MAIL_LIST_DEFAULTS = ["id", "from", "subject", "received"];

export async function mailList(
  args: string[],
  context: MailContext,
): Promise<Record<string, unknown>> {
  const parsed = parseFlags(args, LIST_FLAGS);
  const full = flagBool(parsed, "full", LIST_FLAGS);
  const limit = Math.floor(flagNumber(parsed, "limit", LIST_FLAGS, 20));

  const m365Args = ["outlook", "message", "list"];
  const folder = flagString(parsed, "folder", LIST_FLAGS);
  if (folder !== undefined) {
    m365Args.push(GUID_RE.test(folder) ? "--folderId" : "--folderName", folder);
  }
  const start = flagString(parsed, "start", LIST_FLAGS);
  const end = flagString(parsed, "end", LIST_FLAGS);
  if (start !== undefined) {
    m365Args.push("--startTime", start);
  }
  if (end !== undefined) {
    m365Args.push("--endTime", end);
  }
  m365Args.push(...(await context.m365.userArgs(flagString(parsed, "user", LIST_FLAGS))));

  const items = await context.m365.runJsonArray<MailListItem>(m365Args);
  const shown = items.slice(0, limit);
  const fields = parseFields(flagString(parsed, "fields", LIST_FLAGS), MAIL_LIST_DEFAULTS);

  const rows = shown.map((item) => {
    const row: Record<string, unknown> = {
      id: item.id,
      from: formatFrom(item.from),
      subject: cell(item.subject ?? "", full),
      received: item.receivedDateTime ?? "",
    };
    const extra = fields.filter((f) => !(f in row));
    return { ...row, ...project(item, extra, full) };
  });

  const out: Record<string, unknown> = { mail: rows, count: rows.length };
  if (shown.length < items.length) {
    out.truncated = true;
    out.help = [`Use --limit ${items.length} to see all ${items.length}`];
  }
  return out;
}

export async function mailRead(
  args: string[],
  context: MailContext,
): Promise<Record<string, unknown>> {
  const parsed = parseFlags(args, READ_FLAGS);
  const positionals = parsed.positionals;
  if (positionals.length !== 1) {
    throw new AxiError(
      "mail read expects exactly one message id",
      "VALIDATION_ERROR",
      ["Run `msgraph-axi mail read <id>`"],
    );
  }
  const full = flagBool(parsed, "full", READ_FLAGS);
  const m365Args = [
    "outlook",
    "message",
    "get",
    "--id",
    positionals[0],
    ...(await context.m365.userArgs(flagString(parsed, "user", READ_FLAGS))),
  ];
  const message = await context.m365.runJson<MessageDetail>(m365Args);

  const out: Record<string, unknown> = {
    id: message.id,
    subject: cell(message.subject ?? "", full),
    from: formatFrom(message.from),
    to: formatRecipients(message.toRecipients as EmailAddress[] | undefined),
    received: message.receivedDateTime ?? "",
    snippet: cell(message.bodyPreview ?? "", full),
  };
  if (message.hasAttachments !== undefined) {
    out.hasAttachments = message.hasAttachments;
  }
  if (full && message.body) {
    out.body = message.body.content ?? "";
    out.bodyContentType = message.body.contentType ?? "";
  }
  return out;
}

export async function mailSend(
  args: string[],
  context: MailContext,
): Promise<Record<string, unknown>> {
  const parsed = parseFlags(args, SEND_FLAGS);
  const to = flagString(parsed, "to", SEND_FLAGS);
  const subject = flagString(parsed, "subject", SEND_FLAGS);
  const body = flagString(parsed, "body", SEND_FLAGS);
  const draftId = flagString(parsed, "draft", SEND_FLAGS);
  const execute = flagBool(parsed, "execute", SEND_FLAGS);

  if (draftId !== undefined) {
    if (to !== undefined || subject !== undefined || body !== undefined) {
      throw new AxiError(
        "mail send --draft cannot be combined with --to/--subject/--body",
        "VALIDATION_ERROR",
        ["Just use: msgraph-axi mail send --draft <id> [--execute]"],
      );
    }
    if (!execute) {
      return {
        preview: { draftId, willSend: true },
        execute: false,
        help: ["Run with --execute to send the draft"],
      };
    }
    return mailSendDraft(draftId, flagString(parsed, "user", SEND_FLAGS), context);
  }

  if (to === undefined || subject === undefined || body === undefined) {
    throw new AxiError(
      "mail send requires --to, --subject and --body",
      "VALIDATION_ERROR",
      [
        "Example: msgraph-axi mail send --to a@x.com,b@y.com --subject \"Hi\" --body \"...\"",
      ],
    );
  }

  const preview: Record<string, unknown> = {
    to: to,
    cc: flagString(parsed, "cc", SEND_FLAGS) ?? "",
    bcc: flagString(parsed, "bcc", SEND_FLAGS) ?? "",
    subject: subject,
    bodyChars: body.length,
    bodyPreview: String(cell(body, true)).slice(0, 120),
    attachments: flagString(parsed, "attach", SEND_FLAGS) ?? "",
  };

  if (!execute) {
    return {
      preview,
      execute: false,
      help: ["Run with --execute to send the mail"],
    };
  }
  const m365Args = ["outlook", "mail", "send"];
  m365Args.push("--to", to);
  if (preview.cc) {
    m365Args.push("--cc", preview.cc as string);
  }
  if (preview.bcc) {
    m365Args.push("--bcc", preview.bcc as string);
  }
  m365Args.push("--subject", subject, "--bodyContents", body);
  const bodyType = flagString(parsed, "body-type", SEND_FLAGS);
  if (bodyType !== undefined) {
    m365Args.push("--bodyContentType", bodyType);
  }
  const importance = flagString(parsed, "importance", SEND_FLAGS);
  if (importance !== undefined) {
    m365Args.push("--importance", importance);
  }
  const attach = flagString(parsed, "attach", SEND_FLAGS);
  if (attach !== undefined) {
    m365Args.push("--attachment", attach);
  }
  const mailbox = flagString(parsed, "mailbox", SEND_FLAGS);
  if (mailbox !== undefined) {
    m365Args.push("--mailbox", mailbox);
  }
  const sender = flagString(parsed, "sender", SEND_FLAGS);
  if (sender !== undefined) {
    m365Args.push("--sender", sender);
  }

  await context.m365.run(m365Args);
  return { sent: true, to, subject };
}

export async function mailDelete(
  args: string[],
  context: MailContext,
): Promise<Record<string, unknown>> {
  const parsed = parseFlags(args, DELETE_FLAGS);
  const positionals = parsed.positionals;
  if (positionals.length !== 1) {
    throw new AxiError(
      "mail delete expects exactly one message id",
      "VALIDATION_ERROR",
      ["Run `msgraph-axi mail delete <id>`"],
    );
  }
  const id = positionals[0];
  const execute = flagBool(parsed, "execute", DELETE_FLAGS);
  const confirm = flagString(parsed, "confirm", DELETE_FLAGS);
  if (!execute || confirm !== id) {
    return {
      destructive: true,
      id,
      execute: false,
      help: [
        `Run with --execute --confirm ${id} to delete the message`,
      ],
    };
  }
  const m365Args = [
    "outlook",
    "message",
    "remove",
    "--id",
    id,
    "--force",
    ...(await context.m365.userArgs(flagString(parsed, "user", DELETE_FLAGS))),
  ];
  await context.m365.run(m365Args);
  return { deleted: true, id };
}

const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;