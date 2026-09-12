// Fake m365 CLI for tests. Reads argv, writes fixture JSON on stdout,
// "Error: ..." on stderr with exit 1 for failure modes.
// Set FAKE_M365_LOG=<file> to append every argv as JSON lines for assertions.
import { appendFileSync } from "node:fs";

const args = process.argv.slice(2);
const logFile = process.env.FAKE_M365_LOG;
if (logFile) {
  appendFileSync(logFile, JSON.stringify(args) + "\n");
}

const fail = (message) => {
  process.stderr.write(`Error: ${message}\n`);
  process.exit(1);
};
const print = (value) => {
  process.stdout.write(JSON.stringify(value));
  process.exit(0);
};

const SIZE = 25;
const messages = Array.from({ length: SIZE }, (_, i) => ({
  id: `msg-${i + 1}`,
  subject: `Subject ${i + 1}`,
  from: { emailAddress: { name: "Alice", address: "alice@contoso.com" } },
  toRecipients: [
    { emailAddress: { name: "Bob", address: "bob@contoso.com" } },
  ],
  receivedDateTime: `2026-03-${String((i % 28) + 1).padStart(2, "0")}T09:3${i % 10}:00Z`,
  isRead: i % 2 === 0,
  hasAttachments: i % 5 === 0,
  attachments: i % 5 === 0 ? [{ id: "att-1" }] : [],
  bodyPreview: `Preview of message ${i + 1} `.repeat(30).trim(),
  body: { contentType: "text", content: `Full body of message ${i + 1}` },
}));

const events = Array.from({ length: SIZE }, (_, i) => ({
  id: `evt-${i + 1}`,
  subject: `Event ${i + 1}`,
  start: { dateTime: `2026-03-${String((i % 28) + 1).padStart(2, "0")}T10:00:00`, timeZone: "UTC" },
  end: { dateTime: `2026-03-${String((i % 28) + 1).padStart(2, "0")}T11:00:00`, timeZone: "UTC" },
  location: { displayName: "Room 1" },
}));

const calendars = [
  { id: "cal-1", name: "Calendar", color: "auto" },
  { id: "cal-2", name: "Secondary", color: "auto" },
];

const [cmd, ...rest] = args;

switch (cmd) {
  case "status":
    if (process.env.FAKE_M365_LOGGED_OUT === "1") {
      print("Logged out");
    }
    print({
      connectionName: "contoso",
      connectedAs: "alice@contoso.com",
      authType: "deviceCode",
      appId: "test-app",
      appTenant: "contoso.onmicrosoft.com",
      cloudType: "Public",
    });
    break;

  case "login":
    print({});
    break;

  case "logout":
    if (args.includes("--not-signed-in")) {
      fail("No valid login detected");
    }
    print({});
    break;

  case "outlook":
    switch (rest[0]) {
      case "message":
        if (rest[1] === "list") {
          if (process.env.FAKE_M365_NOT_SIGNED_IN === "1") {
            fail("not signed in");
          }
          print(messages);
        }
        if (rest[1] === "get") {
          const id = rest[rest.indexOf("--id") + 1] ?? "";
          if (id === "missing") fail("ItemNotFound: message not found");
          print(messages.find((m) => m.id === id) ?? messages[0]);
        }
        if (rest[1] === "remove") {
          if (rest.includes("--noconfirm")) fail("missing confirmation");
          print({});
        }
        break;
      case "mail":
        if (rest[1] === "send") {
          if (rest.includes("--noconfirm")) fail("missing to");
          print({});
        }
        break;
      case "calendar":
        if (rest[1] === "list") print(calendars);
        break;
      case "event":
        if (rest[1] === "list") {
          // keep the natural order; agenda sorts client-side
          print([...events].reverse());
        }
        if (rest[1] === "cancel") {
          if (rest.includes("--noconfirm")) fail("missing confirmation");
          print({});
        }
        if (rest[1] === "remove") {
          if (rest.includes("--noconfirm")) fail("missing confirmation");
          print({});
        }
        break;
    }
    break;

  case "request":
    {
      const method = rest[rest.indexOf("--method") + 1] ?? "get";
      const url = rest[rest.indexOf("--url") + 1] ?? "";
      if (args.includes("--graph-fail")) fail("ResourceNotFound");
      if (url.includes("getSchedule")) {
        print([
          {
            scheduleId: "alice@contoso.com",
            availabilityView: "200000020222",
            scheduleItems: [
              { status: "free", start: { dateTime: "2026-03-15T09:00:00" }, end: { dateTime: "2026-03-15T10:00:00" } },
              { status: "busy", start: { dateTime: "2026-03-15T10:00:00" }, end: { dateTime: "2026-03-15T11:00:00" } },
            ],
          },
        ]);
      } else if (method === "get" && url.includes("unreadItemCount")) {
        print({ unreadItemCount: 4 });
      } else if (method === "post" || method === "patch") {
        print({ id: "new-event-id", ...(method === "patch" ? {} : { subject: "x" }) });
      } else {
        print({ value: 1 });
      }
    }
    break;

  default:
    fail(`Unknown command: ${cmd}`);
}