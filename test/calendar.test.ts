import { describe, expect, it } from "vitest";
import { AxiError } from "axi-sdk-js";
import {
  calendarAgenda,
  calendarAvailability,
  calendarCancel,
  calendarCreate,
  calendarDelete,
  calendarList,
  calendarSuggest,
  calendarUpdate,
} from "../src/commands/calendar.js";
import { cleanupContext, expectAxiError, lastM365Call, makeContext, m365Calls, requestBody, type TestBackend } from "./helpers.js";

/** Requests that mutate the mailbox — the read-only lookups do not count. */
function writeCalls(context: TestBackend): string[][] {
  const writeMethods = ["post", "put", "patch", "delete"];
  return m365Calls(context).filter(
    (call) =>
      call[0] === "request" &&
      writeMethods.includes(call[call.indexOf("--method") + 1] ?? ""),
  );
}

describe("calendar list", () => {
  it("lists calendars with id and name", async () => {
    const ctx = makeContext();
    try {
      const out = await calendarList([], ctx);
      expect(out.count).toBe(2);
      expect((out.calendars as Array<Record<string, unknown>>)[0]).toMatchObject({
        id: "cal-1",
        name: "Calendar",
      });
    } finally {
      cleanupContext(ctx);
    }
  });
});

describe("calendar agenda", () => {
  it("defaults to today..+7d, sorts by start and limits to 20", async () => {
    const ctx = makeContext();
    try {
      const out = await calendarAgenda([], ctx);
      expect(out.count).toBe(20);
      expect(out.truncated).toBe(true);
      const rows = out.events as Array<Record<string, unknown>>;
      expect(rows[0].id).toBe("evt-1");
      expect(rows[0].start).toMatch(/^2026-03-\d{2}T10:00:00 UTC$/);
      const call = lastM365Call(ctx);
      expect(call).toEqual(
        expect.arrayContaining([
          "outlook",
          "event",
          "list",
          "--startDateTime",
          expect.stringMatching(/T\d{2}:00:00\.000Z$/),
          "--endDateTime",
          expect.stringMatching(/T\d{2}:00:00\.000Z$/),
          "--userName",
          "alice@contoso.com",
        ]),
      );
    } finally {
      cleanupContext(ctx);
    }
  });

  it("passes calendar, timezone and a custom window through", async () => {
    const ctx = makeContext();
    try {
      await calendarAgenda(
        [
          "--calendar",
          "Secondary",
          "--timezone",
          "W. Europe Standard Time",
          "--start",
          "2026-03-01T00:00:00Z",
          "--end",
          "2026-03-08T00:00:00Z",
        ],
        ctx,
      );
      const call = lastM365Call(ctx);
      expect(call).toEqual(
        expect.arrayContaining([
          "--calendarName",
          "Secondary",
          "--timeZone",
          "W. Europe Standard Time",
        ]),
      );
    } finally {
      cleanupContext(ctx);
    }
  });

  it("supports --fields location", async () => {
    const ctx = makeContext();
    try {
      const out = await calendarAgenda(["--fields", "location", "--limit", "1"], ctx);
      const row = (out.events as Array<Record<string, unknown>>)[0];
      expect(row.location).toBe("Room 1");
    } finally {
      cleanupContext(ctx);
    }
  });
});

describe("calendar create", () => {
  it("dry-runs without --execute", async () => {
    const ctx = makeContext();
    try {
      const out = await calendarCreate(
        ["--subject", "Review", "--start", "2026-03-15T12:00:00", "--end", "2026-03-15T13:00:00"],
        ctx,
      );
      expect(out.execute).toBe(false);
      expect(writeCalls(ctx).length).toBe(0);
    } finally {
      cleanupContext(ctx);
    }
  });

  it("posts to the signed-in user's events with --execute", async () => {
    const ctx = makeContext();
    try {
      const out = await calendarCreate(
        [
          "--subject",
          "Review",
          "--start",
          "2026-03-15T12:00:00",
          "--end",
          "2026-03-15T13:00:00",
          "--location",
          "Room 1",
          "--attendees",
          "bob@contoso.com",
          "--execute",
        ],
        ctx,
      );
      expect(out).toMatchObject({ created: true, id: "new-event-id" });
      const call = lastM365Call(ctx);
      expect(call.slice(0, 3)).toEqual(["request", "--method", "post"]);
      expect(call.join(" ")).toContain(
        "@graph/users/alice@contoso.com/events",
      );
      const body = requestBody(call);
      expect(body.subject).toBe("Review");
      expect((body.start as Record<string, unknown>).timeZone).toBe("UTC");
      expect((body.attendees as Array<Record<string, unknown>>)[0]).toMatchObject({
        emailAddress: { address: "bob@contoso.com" },
      });
    } finally {
      cleanupContext(ctx);
    }
  });

  it("targets a specific calendar id", async () => {
    const ctx = makeContext();
    try {
      await calendarCreate(
        [
          "--subject",
          "X",
          "--start",
          "2026-03-15T12:00:00",
          "--end",
          "2026-03-15T13:00:00",
          "--calendar",
          "cal-1",
          "--execute",
        ],
        ctx,
      );
      const call = lastM365Call(ctx);
      expect(call.join(" ")).toContain("@graph/users/alice@contoso.com/calendars/cal-1/events");
    } finally {
      cleanupContext(ctx);
    }
  });

  it("requires --subject, --start and --end", async () => {
    const ctx = makeContext();
    try {
      await expectAxiError(calendarCreate(["--subject", "X"], ctx), "VALIDATION_ERROR");
    } finally {
      cleanupContext(ctx);
    }
  });
});

describe("calendar update", () => {
  it("patches only the changed fields with --execute", async () => {
    const ctx = makeContext();
    try {
      const out = await calendarUpdate(
        ["evt-1", "--subject", "Renamed", "--execute"],
        ctx,
      );
      expect(out).toMatchObject({ updated: true, id: "evt-1" });
      const call = lastM365Call(ctx);
      expect(call.slice(0, 3)).toEqual(["request", "--method", "patch"]);
      expect(call.join(" ")).toContain("@graph/users/alice@contoso.com/events/evt-1");
      const body = requestBody(call);
      expect(Object.keys(body)).toEqual(["subject"]);
    } finally {
      cleanupContext(ctx);
    }
  });

  it("dry-runs and requires at least one field", async () => {
    const ctx = makeContext();
    try {
      const blocked = await calendarUpdate(["evt-1", "--subject", "Renamed"], ctx);
      expect(blocked.execute).toBe(false);
      await expectAxiError(calendarUpdate(["evt-1"], ctx), "VALIDATION_ERROR");
      expect(writeCalls(ctx).length).toBe(0);
    } finally {
      cleanupContext(ctx);
    }
  });
});

describe("calendar cancel and delete", () => {
  it("both require --execute and an exact --confirm", async () => {
    const ctx = makeContext();
    try {
      const c = await calendarCancel(["evt-1"], ctx);
      expect(c.execute).toBe(false);
      const d = await calendarDelete(["evt-1"], ctx);
      expect(d.execute).toBe(false);
      expect(m365Calls(ctx).length).toBe(0);
    } finally {
      cleanupContext(ctx);
    }
  });

  it("cancels with an optional comment", async () => {
    const ctx = makeContext();
    try {
      const out = await calendarCancel(
        ["evt-1", "--comment", "No longer needed", "--execute", "--confirm", "evt-1"],
        ctx,
      );
      expect(out).toMatchObject({ cancelled: true, id: "evt-1" });
      const call = lastM365Call(ctx);
      expect(call).toEqual(
        expect.arrayContaining([
          "outlook",
          "event",
          "cancel",
          "--id",
          "evt-1",
          "--force",
          "--comment",
          "No longer needed",
        ]),
      );
    } finally {
      cleanupContext(ctx);
    }
  });

  it("deletes with --permanent forwarded", async () => {
    const ctx = makeContext();
    try {
      const out = await calendarDelete(
        ["evt-1", "--permanent", "--execute", "--confirm", "evt-1"],
        ctx,
      );
      expect(out).toMatchObject({ deleted: true, permanent: true });
      const call = lastM365Call(ctx);
      expect(call).toEqual(
        expect.arrayContaining(["outlook", "event", "remove", "--force", "--permanent"]),
      );
    } finally {
      cleanupContext(ctx);
    }
  });
});

describe("calendar availability", () => {
  it("summarizes free/busy per schedule", async () => {
    const ctx = makeContext();
    try {
      const out = await calendarAvailability(
        [
          "--schedules",
          "alice@contoso.com,bob@contoso.com",
          "--start",
          "2026-03-15T09:00:00",
          "--end",
          "2026-03-15T18:00:00",
          "--interval",
          "60",
        ],
        ctx,
      );
      const rows = out.schedules as Array<Record<string, unknown>>;
      expect(rows.length).toBe(1);
      expect(rows[0]).toMatchObject({
        schedule: "alice@contoso.com",
        availabilityView: "200000020222",
        busy: 1,
      });
      expect(rows[0].scheduleItems).toBeUndefined();
      const call = lastM365Call(ctx);
      expect(call.join(" ")).toContain("@graph/users/alice@contoso.com/calendar/getSchedule");
      const body = requestBody(call);
      expect(body.availabilityViewInterval).toBe(60);
      expect((body.schedules as string[]).length).toBe(2);
    } finally {
      cleanupContext(ctx);
    }
  });

  it("expands schedule items with --full", async () => {
    const ctx = makeContext();
    try {
      const out = await calendarAvailability(["--schedules", "a@x.com", "--full"], ctx);
      const rows = out.schedules as Array<Record<string, unknown>>;
      const items = rows[0].scheduleItems as Array<Record<string, unknown>>;
      expect(items[0]).toMatchObject({ status: "free", start: "2026-03-15T09:00:00" });
    } finally {
      cleanupContext(ctx);
    }
  });

  it("requires --schedules", async () => {
    const ctx = makeContext();
    try {
      await expectAxiError(calendarAvailability([], ctx), "VALIDATION_ERROR");
    } finally {
      cleanupContext(ctx);
    }
  });
});

describe("calendar suggest", () => {
  it("finds meeting slots with confidence and availability", async () => {
    const ctx = makeContext();
    try {
      const out = await calendarSuggest(
        ["--attendees", "bob@contoso.com,carol@contoso.com", "--duration", "60"],
        ctx,
      );
      const rows = out.suggestions as Array<Record<string, unknown>>;
      expect(rows.length).toBe(2);
      expect(rows[0]).toMatchObject({
        start: "2026-03-16T09:00:00",
        end: "2026-03-16T10:00:00",
        confidence: "100%",
        available: "2/3",
        organizer: "free",
      });
      expect(rows[0].reason).toBeUndefined();
      const calls = m365Calls(ctx);
      const call = calls[calls.length - 1];
      const body = requestBody(call);
      expect(body.meetingDuration).toBe("PT60M");
      expect((body.attendees as Array<Record<string, unknown>>).length).toBe(2);
      expect(call.join(" ")).toContain("@graph/users/alice@contoso.com/findMeetingTimes");
    } finally {
      cleanupContext(ctx);
    }
  });

  it("expands per-attendee availability with --full", async () => {
    const ctx = makeContext();
    try {
      const out = await calendarSuggest(
        ["--attendees", "bob@contoso.com", "--full"],
        ctx,
      );
      const rows = out.suggestions as Array<Record<string, unknown>>;
      const attendees = rows[0].attendees as Array<Record<string, unknown>>;
      expect(attendees[0]).toMatchObject({
        attendee: "bob@contoso.com",
        availability: "free",
      });
    } finally {
      cleanupContext(ctx);
    }
  });

  it("requires --attendees", async () => {
    const ctx = makeContext();
    try {
      await expectAxiError(calendarSuggest([], ctx), "VALIDATION_ERROR");
    } finally {
      cleanupContext(ctx);
    }
  });
});

describe("calendar time zone handling", () => {
  it("defaults to the mailbox time zone when --timezone is omitted", async () => {
    const ctx = makeContext({ FAKE_M365_TIMEZONE: "Europe/Amsterdam" });
    try {
      await calendarCreate(
        [
          "--subject",
          "Review",
          "--start",
          "2026-03-15T12:00:00",
          "--end",
          "2026-03-15T13:00:00",
          "--execute",
        ],
        ctx,
      );
      const body = requestBody(lastM365Call(ctx));
      expect(body.start).toEqual({
        dateTime: "2026-03-15T12:00:00",
        timeZone: "Europe/Amsterdam",
      });
    } finally {
      cleanupContext(ctx);
    }
  });

  it("prefers an explicit --timezone over the mailbox setting", async () => {
    const ctx = makeContext({ FAKE_M365_TIMEZONE: "Europe/Amsterdam" });
    try {
      const out = await calendarCreate(
        [
          "--subject",
          "Review",
          "--start",
          "2026-03-15T12:00:00",
          "--end",
          "2026-03-15T13:00:00",
          "--timezone",
          "UTC",
        ],
        ctx,
      );
      expect(
        (out.preview as Record<string, unknown>).start,
      ).toEqual({ dateTime: "2026-03-15T12:00:00", timeZone: "UTC" });
    } finally {
      cleanupContext(ctx);
    }
  });

  it("converts a bare local agenda bound for the strict backend", async () => {
    const ctx = makeContext({ FAKE_M365_TIMEZONE: "Europe/Amsterdam" });
    try {
      await calendarAgenda(
        ["--start", "2026-09-15T00:00:00", "--end", "2026-09-16T00:00:00"],
        ctx,
      );
      expect(lastM365Call(ctx)).toEqual(
        expect.arrayContaining([
          "--timeZone",
          "Europe/Amsterdam",
          "--startDateTime",
          "2026-09-14T22:00:00.000Z",
          "--endDateTime",
          "2026-09-15T22:00:00.000Z",
        ]),
      );
    } finally {
      cleanupContext(ctx);
    }
  });

  it("documents the availabilityView codes and the zone it used", async () => {
    const ctx = makeContext({ FAKE_M365_TIMEZONE: "Europe/Amsterdam" });
    try {
      const out = await calendarAvailability(["--schedules", "a@x.com"], ctx);
      expect(out.timezone).toBe("Europe/Amsterdam");
      expect(out.legend).toEqual([
        "0=free",
        "1=tentative",
        "2=busy",
        "3=oof",
        "4=workingElsewhere",
      ]);
      const body = requestBody(lastM365Call(ctx));
      expect(body.startTime).toMatchObject({ timeZone: "Europe/Amsterdam" });
      expect(String((body.startTime as Record<string, unknown>).dateTime)).toMatch(
        /T08:00:00$/,
      );
    } finally {
      cleanupContext(ctx);
    }
  });

  it.each([
    ["15-09-2026"],
    ["2026/09/15"],
    ["next tuesday"],
  ])("rejects the non-ISO bound %s", async (value) => {
    const ctx = makeContext();
    try {
      await expectAxiError(calendarAgenda(["--start", value], ctx), "VALIDATION_ERROR");
      await expectAxiError(
        calendarCreate(
          ["--subject", "X", "--start", value, "--end", "2026-03-15T13:00:00"],
          ctx,
        ),
        "VALIDATION_ERROR",
      );
    } finally {
      cleanupContext(ctx);
    }
  });
});

describe("calendar show-as", () => {
  it("creates an event that reads as free", async () => {
    const ctx = makeContext();
    try {
      const out = await calendarCreate(
        [
          "--subject",
          "Test vanuit msgraph-axi",
          "--start",
          "2026-03-15T13:00:00",
          "--end",
          "2026-03-15T13:05:00",
          "--show-as",
          "free",
          "--execute",
        ],
        ctx,
      );
      expect(out).toMatchObject({ created: true, id: "new-event-id" });
      expect(requestBody(lastM365Call(ctx)).showAs).toBe("free");
    } finally {
      cleanupContext(ctx);
    }
  });

  it("accepts the camelCase alias and rejects unknown states", async () => {
    const ctx = makeContext();
    try {
      const preview = await calendarCreate(
        [
          "--subject",
          "X",
          "--start",
          "2026-03-15T13:00:00",
          "--end",
          "2026-03-15T14:00:00",
          "--showAs",
          "workingElsewhere",
        ],
        ctx,
      );
      expect((preview.preview as Record<string, unknown>).showAs).toBe("workingElsewhere");
      await expectAxiError(
        calendarCreate(
          [
            "--subject",
            "X",
            "--start",
            "2026-03-15T13:00:00",
            "--end",
            "2026-03-15T14:00:00",
            "--show-as",
            "not-a-state",
          ],
          ctx,
        ),
        "VALIDATION_ERROR",
      );
      expect(writeCalls(ctx).length).toBe(0);
    } finally {
      cleanupContext(ctx);
    }
  });

  it("can clear the state on update", async () => {
    const ctx = makeContext();
    try {
      await calendarUpdate(["evt-1", "--show-as", "busy", "--execute"], ctx);
      const body = requestBody(lastM365Call(ctx));
      expect(body).toEqual({ showAs: "busy" });
    } finally {
      cleanupContext(ctx);
    }
  });
});