import { describe, expect, it } from "vitest";
import type { ParsedDateInput } from "../../src/core/date-input.ts";
import { buildCountQuery, buildGraphQuery } from "../../src/core/query-builder.ts";
import type { CountParams, FilterParams, SearchParams } from "../../src/core/schemas.ts";

const params = (over: Partial<SearchParams> = {}): SearchParams => ({
  top: 50,
  ...over,
});

const countParams = (over: Partial<CountParams> = {}): CountParams => ({ ...over }) as FilterParams;

const dt = (iso: string): ParsedDateInput => ({ date: new Date(iso), dateOnly: false });
const dateOnly = (yyyymmdd: string): ParsedDateInput => ({
  date: new Date(`${yyyymmdd}T00:00:00Z`),
  dateOnly: true,
});

describe("buildGraphQuery", () => {
  describe("endpoint", () => {
    it("default /me/messages", () => {
      expect(buildGraphQuery(params()).endpoint).toBe("/me/messages");
    });

    it("folder-scoped when folderId given", () => {
      expect(buildGraphQuery(params(), "fold-1").endpoint).toBe("/me/mailFolders/fold-1/messages");
    });
  });

  describe("baseline (no filters)", () => {
    it("sets $top, $select, $orderby; no $filter, no $search", () => {
      const { query } = buildGraphQuery(params({ top: 25 }));
      expect(query.get("$top")).toBe("25");
      expect(query.get("$select")).toContain("id");
      expect(query.get("$select")).toContain("subject");
      expect(query.get("$select")).toContain("receivedDateTime");
      expect(query.get("$orderby")).toBe("receivedDateTime desc");
      expect(query.has("$filter")).toBe(false);
      expect(query.has("$search")).toBe(false);
    });

    it("$select includes inferenceClassification + parentFolderId", () => {
      const { query } = buildGraphQuery(params());
      expect(query.get("$select")).toContain("inferenceClassification");
      expect(query.get("$select")).toContain("parentFolderId");
    });
  });

  describe("$filter path (no free text)", () => {
    it("since -> receivedDateTime ge ISO", () => {
      const { query } = buildGraphQuery(params({ since: dt("2026-05-01T00:00:00Z") }));
      expect(query.get("$filter")).toBe("receivedDateTime ge 2026-05-01T00:00:00.000Z");
    });

    it("until with explicit datetime -> receivedDateTime le ISO (inclusive instant)", () => {
      const { query } = buildGraphQuery(params({ until: dt("2026-05-09T15:30:00Z") }));
      expect(query.get("$filter")).toBe("receivedDateTime le 2026-05-09T15:30:00.000Z");
    });

    it("until with date-only -> receivedDateTime lt start of next day (whole day included)", () => {
      const { query } = buildGraphQuery(params({ until: dateOnly("2026-05-08") }));
      expect(query.get("$filter")).toBe("receivedDateTime lt 2026-05-09T00:00:00.000Z");
    });

    it("since AND until", () => {
      const { query } = buildGraphQuery(
        params({
          since: dt("2026-05-01T00:00:00Z"),
          until: dateOnly("2026-05-08"),
        }),
      );
      expect(query.get("$filter")).toBe(
        "receivedDateTime ge 2026-05-01T00:00:00.000Z and receivedDateTime lt 2026-05-09T00:00:00.000Z",
      );
    });

    it("is_unread true -> isRead eq false", () => {
      expect(buildGraphQuery(params({ is_unread: true })).query.get("$filter")).toBe(
        "isRead eq false",
      );
    });

    it("is_unread false -> isRead eq true", () => {
      expect(buildGraphQuery(params({ is_unread: false })).query.get("$filter")).toBe(
        "isRead eq true",
      );
    });

    it("importance -> importance eq 'value'", () => {
      expect(buildGraphQuery(params({ importance: "high" })).query.get("$filter")).toBe(
        "importance eq 'high'",
      );
    });

    it("has_attachment -> hasAttachments eq true", () => {
      expect(buildGraphQuery(params({ has_attachment: true })).query.get("$filter")).toBe(
        "hasAttachments eq true",
      );
    });

    it("inference_classification -> inferenceClassification eq 'focused'", () => {
      expect(
        buildGraphQuery(params({ inference_classification: "focused" })).query.get("$filter"),
      ).toBe("inferenceClassification eq 'focused'");
    });

    it("inference_classification 'other' filters as 'other'", () => {
      expect(
        buildGraphQuery(params({ inference_classification: "other" })).query.get("$filter"),
      ).toBe("inferenceClassification eq 'other'");
    });

    it("multiple structured filters AND-joined", () => {
      const { query } = buildGraphQuery(
        params({
          has_attachment: true,
          importance: "high",
          is_unread: true,
          since: dt("2026-05-01T00:00:00Z"),
        }),
      );
      expect(query.get("$filter")).toBe(
        "receivedDateTime ge 2026-05-01T00:00:00.000Z and isRead eq false and hasAttachments eq true and importance eq 'high'",
      );
    });

    it("keeps $orderby in $filter path", () => {
      const { query } = buildGraphQuery(params({ has_attachment: true }));
      expect(query.get("$orderby")).toBe("receivedDateTime desc");
    });
  });

  describe("$search path (any text/people field present)", () => {
    it("query alone -> $search 'query'", () => {
      const { query } = buildGraphQuery(params({ query: "interview" }));
      expect(query.get("$search")).toBe('"interview"');
      expect(query.has("$filter")).toBe(false);
      expect(query.has("$orderby")).toBe(false);
    });

    it("from -> from:value", () => {
      expect(buildGraphQuery(params({ from: "goldman.com" })).query.get("$search")).toBe(
        '"from:goldman.com"',
      );
    });

    it("to -> to:value", () => {
      expect(buildGraphQuery(params({ to: "me@x.com" })).query.get("$search")).toBe(
        '"to:me@x.com"',
      );
    });

    it("subject_contains -> subject:value", () => {
      expect(buildGraphQuery(params({ subject_contains: "assessment" })).query.get("$search")).toBe(
        '"subject:assessment"',
      );
    });

    it("body_contains -> body:value", () => {
      expect(buildGraphQuery(params({ body_contains: "schedule" })).query.get("$search")).toBe(
        '"body:schedule"',
      );
    });

    it("multiple fields space-joined; query first", () => {
      const { query } = buildGraphQuery(
        params({ query: "interview", from: "goldman.com", subject_contains: "assessment" }),
      );
      expect(query.get("$search")).toBe('"interview from:goldman.com subject:assessment"');
    });

    it("structured booleans fold into KQL when text present", () => {
      const { query } = buildGraphQuery(
        params({ query: "interview", has_attachment: true, is_unread: true, importance: "high" }),
      );
      const s = query.get("$search");
      expect(s).toContain("interview");
      expect(s).toContain("hasattachment:yes");
      expect(s).toContain("read:no");
      expect(s).toContain("importance:high");
    });

    it("until in KQL is exclusive against next day (whole day included)", () => {
      const { query } = buildGraphQuery(
        params({
          query: "interview",
          since: dateOnly("2026-05-01"),
          until: dateOnly("2026-05-08"),
        }),
      );
      const s = query.get("$search") ?? "";
      expect(s).toContain("received>=2026-05-01");
      expect(s).toContain("received<2026-05-09");
    });

    it("$filter omitted in search path", () => {
      const { query } = buildGraphQuery(params({ query: "x", has_attachment: true }));
      expect(query.has("$filter")).toBe(false);
    });
  });

  describe("inference_classification + text search", () => {
    it("throws when combined with query (KQL has no inference keyword)", () => {
      expect(() =>
        buildGraphQuery(params({ query: "interview", inference_classification: "focused" })),
      ).toThrow(/inference_classification/);
    });

    it("throws when combined with from", () => {
      expect(() =>
        buildGraphQuery(params({ from: "x@y.com", inference_classification: "focused" })),
      ).toThrow(/inference_classification/);
    });
  });

  describe("OData string escaping", () => {
    it("doubles single quotes in folder ID is not relevant here — enums only need exact match", () => {
      expect(buildGraphQuery(params({ importance: "low" })).query.get("$filter")).toBe(
        "importance eq 'low'",
      );
    });
  });
});

describe("buildCountQuery", () => {
  it("sets $count=true and minimal $top/$select", () => {
    const { endpoint, query } = buildCountQuery(countParams());
    expect(endpoint).toBe("/me/messages");
    expect(query.get("$count")).toBe("true");
    expect(query.get("$top")).toBe("1");
    expect(query.get("$select")).toBe("id");
  });

  it("does NOT set $orderby (Graph rejects $count + $orderby)", () => {
    const { query } = buildCountQuery(countParams({ since: dt("2026-05-01T00:00:00Z") }));
    expect(query.has("$orderby")).toBe(false);
  });

  it("applies $filter for structured path", () => {
    const { query } = buildCountQuery(countParams({ is_unread: true }));
    expect(query.get("$filter")).toBe("isRead eq false");
  });

  it("applies $search for text path", () => {
    const { query } = buildCountQuery(countParams({ from: "goldman.com" }));
    expect(query.get("$search")).toBe('"from:goldman.com"');
    expect(query.has("$filter")).toBe(false);
  });

  it("folder-scoped endpoint", () => {
    expect(buildCountQuery(countParams(), "fold-1").endpoint).toBe(
      "/me/mailFolders/fold-1/messages",
    );
  });

  it("throws on inference_classification + text combo", () => {
    expect(() =>
      buildCountQuery(countParams({ query: "x", inference_classification: "focused" })),
    ).toThrow(/inference_classification/);
  });
});
