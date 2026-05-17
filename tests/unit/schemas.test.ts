import { describe, expect, it } from "vitest";
import { FolderSchema, LeanMessageSchema, SearchParamsSchema } from "../../src/core/schemas.ts";

describe("SearchParamsSchema", () => {
  it("applies default top=50 when omitted", () => {
    const parsed = SearchParamsSchema.parse({});
    expect(parsed.top).toBe(50);
  });

  it("respects explicit top within range", () => {
    expect(SearchParamsSchema.parse({ top: 200 }).top).toBe(200);
  });

  it("rejects top > 500", () => {
    expect(() => SearchParamsSchema.parse({ top: 1000 })).toThrow();
  });

  it("rejects top < 1", () => {
    expect(() => SearchParamsSchema.parse({ top: 0 })).toThrow();
  });

  it("rejects non-integer top", () => {
    expect(() => SearchParamsSchema.parse({ top: 1.5 })).toThrow();
  });

  it("transforms since string into ParsedDateInput (date-only)", () => {
    const parsed = SearchParamsSchema.parse({ since: "2026-05-01" });
    expect(parsed.since?.date).toBeInstanceOf(Date);
    expect(parsed.since?.date.toISOString()).toBe("2026-05-01T00:00:00.000Z");
    expect(parsed.since?.dateOnly).toBe(true);
  });

  it("transforms relative since (-7d) into ParsedDateInput (dateOnly=false)", () => {
    const parsed = SearchParamsSchema.parse({ since: "-7d" });
    expect(parsed.since?.date).toBeInstanceOf(Date);
    expect(parsed.since?.dateOnly).toBe(false);
  });

  it("rejects invalid since string", () => {
    expect(() => SearchParamsSchema.parse({ since: "not-a-date" })).toThrow();
  });

  it("accepts valid importance values", () => {
    for (const i of ["low", "normal", "high"] as const) {
      expect(SearchParamsSchema.parse({ importance: i }).importance).toBe(i);
    }
  });

  it("rejects invalid importance", () => {
    expect(() => SearchParamsSchema.parse({ importance: "urgent" })).toThrow();
  });

  it("accepts inference_classification focused/other", () => {
    expect(
      SearchParamsSchema.parse({ inference_classification: "focused" }).inference_classification,
    ).toBe("focused");
    expect(
      SearchParamsSchema.parse({ inference_classification: "other" }).inference_classification,
    ).toBe("other");
  });

  it("rejects invalid inference_classification", () => {
    expect(() => SearchParamsSchema.parse({ inference_classification: "junk" })).toThrow();
  });

  it("rejects unknown fields", () => {
    expect(() => SearchParamsSchema.parse({ random_field: "x" })).toThrow();
  });

  it("accepts a fully populated query", () => {
    const parsed = SearchParamsSchema.parse({
      query: "interview",
      from: "goldman.com",
      to: "me@x.com",
      subject_contains: "assessment",
      body_contains: "schedule",
      since: "-30d",
      until: "2026-05-09",
      has_attachment: true,
      folder: "Jobs",
      is_unread: false,
      importance: "high",
      top: 100,
    });
    expect(parsed.top).toBe(100);
    expect(parsed.since?.date).toBeInstanceOf(Date);
    expect(parsed.until?.date).toBeInstanceOf(Date);
  });
});

describe("LeanMessageSchema", () => {
  it("accepts a valid lean message", () => {
    const m = LeanMessageSchema.parse({
      id: "abc",
      from: { name: "Goldman Recruiter", address: "r@goldman.com" },
      subject: "Assessment invitation",
      received_at: "2026-05-08T14:00:00Z",
      has_attachment: false,
      is_read: false,
      conversation_id: "conv-1",
      web_link: "https://outlook.office.com/mail/inbox/id/abc",
      body_preview: "We are pleased to invite you...",
    });
    expect(m.id).toBe("abc");
  });

  it("from.name optional, from.address required", () => {
    expect(() =>
      LeanMessageSchema.parse({
        id: "abc",
        from: { name: "X" },
        subject: "x",
        received_at: "2026-05-08T14:00:00Z",
        has_attachment: false,
        is_read: false,
        conversation_id: "c",
        web_link: "https://x.com",
        body_preview: "",
      }),
    ).toThrow();
  });
});

describe("FolderSchema", () => {
  it("accepts a folder", () => {
    const f = FolderSchema.parse({
      id: "fold-1",
      display_name: "Inbox",
      total_item_count: 100,
      unread_item_count: 5,
    });
    expect(f.display_name).toBe("Inbox");
  });

  it("accepts optional parent_folder_id", () => {
    const f = FolderSchema.parse({
      id: "fold-2",
      display_name: "Jobs",
      parent_folder_id: "fold-1",
      total_item_count: 10,
      unread_item_count: 0,
    });
    expect(f.parent_folder_id).toBe("fold-1");
  });
});
