import { describe, expect, it } from "vitest";
import { parseDateInput } from "../../src/core/date-input.ts";

const NOW = new Date("2026-05-09T12:00:00Z");

describe("parseDateInput", () => {
  describe("ISO 8601", () => {
    it("parses date-only ISO as UTC midnight, dateOnly=true", () => {
      const r = parseDateInput("2026-05-01", NOW);
      expect(r.date.toISOString()).toBe("2026-05-01T00:00:00.000Z");
      expect(r.dateOnly).toBe(true);
    });

    it("parses full ISO datetime, dateOnly=false", () => {
      const r = parseDateInput("2026-05-01T10:30:00Z", NOW);
      expect(r.date.toISOString()).toBe("2026-05-01T10:30:00.000Z");
      expect(r.dateOnly).toBe(false);
    });

    it("parses ISO with offset, dateOnly=false", () => {
      const r = parseDateInput("2026-05-01T10:00:00+02:00", NOW);
      expect(r.date.toISOString()).toBe("2026-05-01T08:00:00.000Z");
      expect(r.dateOnly).toBe(false);
    });
  });

  describe("relative shorthand", () => {
    it("parses -7d as 7 days before now, dateOnly=false", () => {
      const r = parseDateInput("-7d", NOW);
      expect(r.date.toISOString()).toBe("2026-05-02T12:00:00.000Z");
      expect(r.dateOnly).toBe(false);
    });

    it("parses -2w as 14 days before now", () => {
      expect(parseDateInput("-2w", NOW).date.toISOString()).toBe("2026-04-25T12:00:00.000Z");
    });

    it("parses -3h as 3 hours before now", () => {
      expect(parseDateInput("-3h", NOW).date.toISOString()).toBe("2026-05-09T09:00:00.000Z");
    });

    it("parses -30m as 30 minutes before now", () => {
      expect(parseDateInput("-30m", NOW).date.toISOString()).toBe("2026-05-09T11:30:00.000Z");
    });

    it("parses -1d (single digit)", () => {
      expect(parseDateInput("-1d", NOW).date.toISOString()).toBe("2026-05-08T12:00:00.000Z");
    });

    it("parses -365d (multi-digit)", () => {
      expect(parseDateInput("-365d", NOW).date.toISOString()).toBe("2025-05-09T12:00:00.000Z");
    });
  });

  describe("invalid input", () => {
    it("throws on empty string", () => {
      expect(() => parseDateInput("", NOW)).toThrow();
    });

    it("throws on garbage", () => {
      expect(() => parseDateInput("not-a-date", NOW)).toThrow();
    });

    it("throws on positive offset (+7d)", () => {
      expect(() => parseDateInput("+7d", NOW)).toThrow();
    });

    it("throws on bare number", () => {
      expect(() => parseDateInput("7", NOW)).toThrow();
    });

    it("throws on unknown unit (-7y)", () => {
      expect(() => parseDateInput("-7y", NOW)).toThrow();
    });

    it("throws on invalid ISO (2026-13-01)", () => {
      expect(() => parseDateInput("2026-13-01", NOW)).toThrow();
    });
  });
});
