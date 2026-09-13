import { AxiError } from "axi-sdk-js";
import { readFileSync, writeFileSync } from "node:fs";
import { basename, extname, join, resolve } from "node:path";
import { PnpCliBackend } from "../backend.js";
import {
  getGraphValue,
  graphPrefix,
  odataValue,
  searchQuery,
  type GraphContext,
} from "../graph.js";
import {
  parseFlags,
  flagString,
  flagNumber,
  flagBool,
  type FlagDef,
} from "../flags.js";
import { cell, formatFrom, parseFields } from "../toon.js";
import type { EmailAddress } from "../toon.js";

export interface MailFlowContext extends GraphContext {
  m365: PnpCliBackend;
  user?: string;
}

interface SearchResult {
  id?: string;
  subject?: string;
  from?: EmailAddress;
  receivedDateTime?: string;
  conversationId?: string;
  hasAttachments?: boolean;
  [key: string]: unknown;
}

interface DraftMessage extends SearchResult {
  toRecipients?: EmailAddress[];
  ccRecipients?: EmailAddress[];
  bodyPreview?: string;
  body?: { contentType?: string; content?: string };
}

const SEARCH_FLAGS: Record<string, FlagDef> = {
  user: { type: "string" },
  search: { type: "string" },
  limit: { type: "number" },
  fields: { type: "string" },
  full: { type: "boolean" },
};

const DRAFT_FLAGS: Record<string, FlagDef> = {
  to: { type: "string" },
  cc: { type: "string" },
  bcc: { type: "string" },
  subject: { type: "string", aliases: ["s"] },
  body: { type: "string" },
  "body-type": { type: "string" },
  importance: { type: "string" },
  user: { type: "string" },
  execute: { type: "boolean" },
};

const THREAD_FLAGS: Record<string, FlagDef> = {
  user: { type: "string" },
  limit: { type: "number" },
  full: { type: "boolean" },
};

const ATTACHMENT_FLAGS: Record<string, FlagDef> = {
  message: { type: "string" },
  out: { type: "string" },
};

const SEARCH_DEFAULTS = ["id", "from", "subject", "received"];

export async function mailSearch(
  args: string[],
  context: MailFlowContext,
): Promise<Record<string, unknown>> {
  const parsed = parseFlags(args, SEARCH_FLAGS);
  const query = flagString(parsed, "search", SEARCH_FLAGS);
  if (query === undefined || query.trim() === "") {
    throw new AxiError(
      "mail search requires --search <query>",
      "VALIDATION_ERROR",
      ['Example: msgraph-axi mail search --search "invoice overdue"'],
    );
  }
  const full = flagBool(parsed, "full", SEARCH_FLAGS);
  const limit = Math.floor(flagNumber(parsed, "limit", SEARCH_FLAGS, 20));
  const items = await getGraphValue<SearchResult>(
    context,
    flagString(parsed, "user", SEARCH_FLAGS),
    `/messages${searchQuery(query, limit, [
      "id",
      "subject",
      "from",
      "receivedDateTime",
      "bodyPreview",
      "conversationId",
      "hasAttachments",
    ])}`,
  );
  const fields = parseFields(
    flagString(parsed, "fields", SEARCH_FLAGS),
    SEARCH_DEFAULTS,
  );
  const rows = items.map((item) => {
    const row: Record<string, unknown> = {
      id: item.id,
      from: formatFrom(item.from),
      subject: cell(item.subject ?? "", full),
      received: item.receivedDateTime ?? "",
    };
    const extra = fields.filter((f) => !(f in row));
    for (const field of extra) {
      row[field] = cell(
        field === "from" ? formatFrom(item.from) : item[field],
        full,
      );
    }
    return row;
  });
  const out: Record<string, unknown> = {
    mail: rows,
    count: rows.length,
    query,
  };
  if (rows.length === limit) {
    out.truncated = true;
    out.help = [`Use --limit ${limit * 2} for more results`];
  }
  return out;
}

export interface DraftInput {
  to: string;
  subject: string;
  body: string;
  bodyType?: string;
  cc?: string;
  bcc?: string;
  importance?: string;
  /** Sets the Graph `from` address (m365 `--sender`). */
  sender?: string;
  attachments?: Array<Record<string, string>>;
  userFlag?: string;
}

/** Graph accepts up to 3 MB of file attachments on a message. */
const MAX_ATTACHMENT_BYTES = 3 * 1024 * 1024;

const MIME_BY_EXTENSION: Record<string, string> = {
  pdf: "application/pdf",
  txt: "text/plain",
  csv: "text/csv",
  json: "application/json",
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  gif: "image/gif",
  zip: "application/zip",
  doc: "application/msword",
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  xls: "application/vnd.ms-excel",
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  ppt: "application/vnd.ms-powerpoint",
  pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
};

/** Read comma-separated `--attach` paths as Graph file attachments. */
export function readAttachments(paths: string): Array<Record<string, string>> {
  return paths
    .split(",")
    .map((path) => path.trim())
    .filter(Boolean)
    .map((filePath) => {
      let data: Buffer;
      try {
        data = readFileSync(filePath);
      } catch {
        throw new AxiError(
          `Attachment not found: ${filePath}`,
          "VALIDATION_ERROR",
          ["Pass comma-separated paths: --attach report.pdf,sheet.xlsx"],
        );
      }
      if (data.byteLength > MAX_ATTACHMENT_BYTES) {
        throw new AxiError(
          `Attachment ${basename(filePath)} is ${(data.byteLength / 1024 / 1024).toFixed(1)} MB; a draft carries at most 3 MB`,
          "VALIDATION_ERROR",
          [
            "Attach bigger files in Outlook after the draft is created",
            "Or share a link instead of the file",
          ],
        );
      }
      const extension = extname(filePath).slice(1).toLowerCase();
      return {
        "@odata.type": "#microsoft.graph.fileAttachment",
        name: basename(filePath),
        contentType: MIME_BY_EXTENSION[extension] ?? "application/octet-stream",
        contentBytes: data.toString("base64"),
      };
    });
}

/**
 * Create a draft message through the Graph JSON bridge. The body travels as
 * JSON, so line breaks stay escaped and multi-line mail cannot lose content on
 * the way to the backend.
 */
export async function createDraftMessage(
  context: MailFlowContext,
  input: DraftInput,
): Promise<string> {
  const payload: Record<string, unknown> = {
    subject: input.subject,
    body: { contentType: input.bodyType ?? "text", content: input.body },
    toRecipients: recipients(input.to),
    isDraft: true,
  };
  if (input.cc !== undefined) {
    payload.ccRecipients = recipients(input.cc);
  }
  if (input.bcc !== undefined) {
    payload.bccRecipients = recipients(input.bcc);
  }
  if (input.importance !== undefined) {
    payload.importance = input.importance;
  }
  if (input.sender !== undefined) {
    payload.from = recipients(input.sender)[0];
  }
  if (input.attachments !== undefined && input.attachments.length > 0) {
    payload.attachments = input.attachments;
  }
  const prefix = await graphPrefix(context, input.userFlag);
  const draft = await context.m365.request<DraftMessage>(
    "post",
    `${prefix}/messages`,
    payload,
  );
  return draft.id ?? "";
}

export async function mailDraft(
  args: string[],
  context: MailFlowContext,
): Promise<Record<string, unknown>> {
  const parsed = parseFlags(args, DRAFT_FLAGS);
  const to = flagString(parsed, "to", DRAFT_FLAGS);
  const subject = flagString(parsed, "subject", DRAFT_FLAGS);
  const body = flagString(parsed, "body", DRAFT_FLAGS);
  if (to === undefined || subject === undefined || body === undefined) {
    throw new AxiError(
      "mail draft requires --to, --subject and --body",
      "VALIDATION_ERROR",
      [
        'Example: msgraph-axi mail draft --to a@x.com --subject "Hi" --body "..."',
      ],
    );
  }
  const bodyType = flagString(parsed, "body-type", DRAFT_FLAGS) ?? "text";
  const preview: Record<string, unknown> = {
    to,
    subject,
    bodyChars: body.length,
    bodyType,
  };
  if (!flagBool(parsed, "execute", DRAFT_FLAGS)) {
    return {
      preview,
      execute: false,
      help: ["Run with --execute to save the draft"],
    };
  }
  const payloadId = await createDraftMessage(context, {
    to,
    subject,
    body,
    bodyType,
    cc: flagString(parsed, "cc", DRAFT_FLAGS),
    bcc: flagString(parsed, "bcc", DRAFT_FLAGS),
    importance: flagString(parsed, "importance", DRAFT_FLAGS),
    userFlag: flagString(parsed, "user", DRAFT_FLAGS),
  });
  return { draft: true, id: payloadId, subject, saveToSent: false };
}

export async function mailSendDraft(
  draftId: string,
  userFlag: string | undefined,
  context: MailFlowContext,
): Promise<Record<string, unknown>> {
  const prefix = await graphPrefix(context, userFlag);
  await context.m365.run([
    "request",
    "--method",
    "post",
    "--url",
    `${prefix}/messages/${draftId}/send`,
  ]);
  return { sent: true, draftId };
}

interface ThreadMessage extends SearchResult {
  bodyPreview?: string;
  body?: { content?: string };
  toRecipients?: EmailAddress[];
}

export async function mailThread(
  args: string[],
  context: MailFlowContext,
): Promise<Record<string, unknown>> {
  const parsed = parseFlags(args, THREAD_FLAGS);
  const positionals = parsed.positionals;
  if (positionals.length !== 1) {
    throw new AxiError(
      "mail thread expects exactly one conversation id",
      "VALIDATION_ERROR",
      ["Run `msgraph-axi mail thread <conversationId>`"],
    );
  }
  const full = flagBool(parsed, "full", THREAD_FLAGS);
  const limit = Math.floor(flagNumber(parsed, "limit", THREAD_FLAGS, 100));
  const items = await getGraphValue<ThreadMessage>(
    context,
    flagString(parsed, "user", THREAD_FLAGS),
    `/messages?$filter=conversationId%20eq%20'${odataValue(
      positionals[0],
    )}'&$top=${limit}&$select=id,subject,from,toRecipients,receivedDateTime,bodyPreview,body`,
  );
  const sorted = [...items].sort((a, b) =>
    (a.receivedDateTime ?? "").localeCompare(b.receivedDateTime ?? ""),
  );
  const rows = sorted.map((message) => {
    const row: Record<string, unknown> = {
      id: message.id,
      from: formatFrom(message.from),
      received: message.receivedDateTime ?? "",
      snippet: cell(message.bodyPreview ?? "", full),
    };
    if (full && message.body?.content !== undefined) {
      row.bodyContent = cell(message.body.content, true);
    }
    return row;
  });
  return {
    thread: rows,
    count: rows.length,
    conversationId: positionals[0],
    subject: cell(sorted[0]?.subject ?? "", false),
  };
}

export async function mailAttachmentGet(
  args: string[],
  context: MailFlowContext,
): Promise<Record<string, unknown>> {
  const parsed = parseFlags(args, ATTACHMENT_FLAGS);
  const positionals = parsed.positionals;
  if (positionals.length !== 1) {
    throw new AxiError(
      "mail attachment get expects exactly one attachment id",
      "VALIDATION_ERROR",
      [
        "Run `msgraph-axi mail attachment get <attachmentId> --message <messageId>`",
      ],
    );
  }
  const messageId = flagString(parsed, "message", ATTACHMENT_FLAGS);
  if (messageId === undefined) {
    throw new AxiError(
      "mail attachment get requires --message <messageId>",
      "VALIDATION_ERROR",
      ["Add --message <id> from `mail list` or `mail read`"],
    );
  }
  const attachmentId = positionals[0];
  const prefix = await graphPrefix(context, undefined);
  const base = `${prefix}/messages/${messageId}/attachments/${attachmentId}`;
  const meta = await context.m365.runJson<{
    name?: string;
    contentType?: string;
    size?: number;
  }>(["request", "--method", "get", "--url", base]);
  const name = meta.name ?? attachmentId;
  const out = resolve(
    flagString(parsed, "out", ATTACHMENT_FLAGS) ?? join(process.cwd(), name),
  );
  const downloaded = await context.m365.run([
    "request",
    "--method",
    "get",
    "--url",
    `${base}/$value`,
    "--filePath",
    out,
  ]);
  void downloaded;
  return {
    saved: out,
    attachment: name,
    contentType: meta.contentType ?? "",
    size: meta.size ?? 0,
  };
}

function recipients(value: string): Array<{ emailAddress: { address: string } }> {
  return value
    .split(",")
    .map((a) => a.trim())
    .filter(Boolean)
    .map((address) => ({ emailAddress: { address } }));
}