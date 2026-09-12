import { describe, expect, it } from "vitest";
import { AxiError } from "axi-sdk-js";
import {
  calendarAgenda,
  calendarAvailability,
  calendarCancel,
  calendarCreate,
  calendarDelete,
  calendarList,
  calendarUpdate,
} from "../src/commands/calendar.js";
import { cleanupContext, expectAxiError, lastM365Call, makeContext, m365Calls } from "./helpers.js";

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
      expect(m365Calls(ctx).length).toBe(0);
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
      const body = JSON.parse(call[call.indexOf("--body") + 1] ?? "{}") as Record<string, unknown>;
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
      const body = JSON.parse(call[call.indexOf("--body") + 1] ?? "{}") as Record<string, unknown>;
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
      expect(m365Calls(ctx).length).toBe(0);
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
      const body = JSON.parse(call[call.indexOf("--body") + 1] ?? "{}") as Record<string, unknown>;
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