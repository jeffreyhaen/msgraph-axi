import { PnpCliBackend } from "../backend.js";
import { localDayIso } from "../toon.js";
import type { MailContext, MailListItem } from "./mail.js";

export interface HomeContext extends MailContext {
  m365: PnpCliBackend;
}

interface InboxCount {
  unreadItemCount?: number;
}

/**
 * Content-first home view: sign-in status plus the two numbers an agent needs
 * to act immediately — unread mail and today's agenda. Every part degrades to
 * an actionable auth hint instead of failing the whole view.
 */
export async function home(
  args: string[],
  context: HomeContext,
): Promise<Record<string, unknown>> {
  const status = await context.m365.status();
  if (!status) {
    return {
      signedIn: false,
      help: [
        "Run `msgraph-axi auth login` to sign in",
        "Then `msgraph-axi mail list` or `msgraph-axi calendar agenda`",
      ],
    };
  }

  const out: Record<string, unknown> = {
    signedIn: true,
    connectedAs: status.connectedAs ?? "",
  };

  try {
    const inbox = await context.m365.runJson<InboxCount>([
      "request",
      "--method",
      "get",
      "--url",
      "@graph/me/mailFolders/inbox?$select=unreadItemCount",
    ]);
    out.unreadMail = inbox.unreadItemCount ?? 0;
  } catch {
    out.unreadMail = undefined;
  }

  try {
    const events = await context.m365.runJsonArray<Record<string, unknown>>([
      "outlook",
      "event",
      "list",
      "--startDateTime",
      localDayIso(0),
      "--endDateTime",
      localDayIso(1),
      "--userName",
      status.connectedAs ?? "",
    ]);
    out.today = events.slice(0, 3).map((event) => ({
      id: event.id,
      subject: event.subject,
      start: (event.start as { dateTime?: string } | undefined)?.dateTime ?? "",
    }));
  } catch {
    out.today = [];
  }

  out.help = [
    "Run `msgraph-axi mail list` for mail",
    "Run `msgraph-axi calendar agenda` for the coming week",
  ];
  return out;
}