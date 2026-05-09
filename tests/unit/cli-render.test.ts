import { describe, expect, it } from "vitest";
import { renderFolders, renderFullMessage, renderSearchResults } from "../../src/cli-render.ts";
import type { Folder, FullMessage, LeanMessage, SearchResult } from "../../src/core/schemas.ts";

const lean = (over: Partial<LeanMessage> = {}): LeanMessage => ({
  id: "abc",
  from: { name: "Example Recruiter", address: "recruiter@example.com" },
  subject: "Assessment invitation",
  received_at: "2026-05-08T14:43:14Z",
  has_attachment: false,
  is_read: false,
  conversation_id: "conv-1",
  web_link: "https://outlook.office.com/x",
  body_preview: "Dear candidate...",
  ...over,
});

describe("renderSearchResults", () => {
  it("renders empty result", () => {
    const out = renderSearchResults({ results: [], total_returned: 0, more_available: false });
    expect(out).toContain("0 results");
  });

  it("includes from address, subject, preview, and date", () => {
    const result: SearchResult = {
      results: [lean()],
      total_returned: 1,
      more_available: false,
    };
    const out = renderSearchResults(result);
    expect(out).toContain("recruiter@example.com");
    expect(out).toContain("Assessment invitation");
    expect(out).toContain("Dear candidate");
    expect(out).toContain("2026-05-08");
  });

  it("indicates unread messages", () => {
    const out = renderSearchResults({
      results: [lean({ is_read: false })],
      total_returned: 1,
      more_available: false,
    });
    expect(out).toMatch(/unread|●|\*/i);
  });

  it("indicates attachments", () => {
    const out = renderSearchResults({
      results: [lean({ has_attachment: true })],
      total_returned: 1,
      more_available: false,
    });
    expect(out).toMatch(/attachment|📎|\[a\]|@/i);
  });

  it("shows more_available footer", () => {
    const out = renderSearchResults({
      results: [lean()],
      total_returned: 1,
      more_available: true,
    });
    expect(out).toMatch(/more available|more results/i);
  });

  it("falls back to address when name missing", () => {
    const out = renderSearchResults({
      results: [lean({ from: { address: "x@y.com" } })],
      total_returned: 1,
      more_available: false,
    });
    expect(out).toContain("x@y.com");
  });
});

describe("renderFullMessage", () => {
  const full: FullMessage = {
    ...lean(),
    to: [{ name: "Harry", address: "me@x.com" }],
    cc: [],
    body: "Plain body content here.",
    body_content_type: "text",
    importance: "high",
    internet_message_headers: [],
  };

  it("includes from, to, subject, date, body", () => {
    const out = renderFullMessage(full);
    expect(out).toContain("recruiter@example.com");
    expect(out).toContain("me@x.com");
    expect(out).toContain("Assessment invitation");
    expect(out).toContain("Plain body content here.");
    expect(out).toContain("2026-05-08");
  });

  it("flags importance when not normal", () => {
    const out = renderFullMessage(full);
    expect(out).toMatch(/high/i);
  });

  it("does not flag importance when normal", () => {
    const out = renderFullMessage({ ...full, importance: "normal" });
    expect(out).not.toMatch(/^Importance:/m);
  });

  it("renders cc when present", () => {
    const out = renderFullMessage({
      ...full,
      cc: [{ address: "cc@x.com" }],
    });
    expect(out).toContain("cc@x.com");
  });

  it("notes html content type warning", () => {
    const out = renderFullMessage({ ...full, body_content_type: "html", body: "<p>x</p>" });
    expect(out).toMatch(/html/i);
  });

  it("renders Authentication-Results header when present", () => {
    const out = renderFullMessage({
      ...full,
      internet_message_headers: [
        { name: "Authentication-Results", value: "spf=pass; dkim=pass; dmarc=pass" },
      ],
    });
    expect(out).toContain("Authentication-Results");
    expect(out).toContain("spf=pass; dkim=pass; dmarc=pass");
  });

  it("renders Return-Path header when present", () => {
    const out = renderFullMessage({
      ...full,
      internet_message_headers: [{ name: "Return-Path", value: "<bounces@example.com>" }],
    });
    expect(out).toContain("Return-Path");
    expect(out).toContain("bounces@example.com");
  });

  it("does not render auth section when no relevant headers", () => {
    const out = renderFullMessage({
      ...full,
      internet_message_headers: [{ name: "X-Random-Header", value: "noise" }],
    });
    expect(out).not.toContain("Authentication-Results");
    expect(out).not.toContain("Return-Path");
  });

  it("flags inference_classification when 'other'", () => {
    const out = renderFullMessage({ ...full, inference_classification: "other" });
    expect(out).toMatch(/other/i);
  });

  it("does not flag inference_classification when 'focused'", () => {
    const out = renderFullMessage({ ...full, inference_classification: "focused" });
    expect(out).not.toMatch(/^Classification:/m);
  });

  it("does not flag inference_classification when absent", () => {
    const out = renderFullMessage(full);
    expect(out).not.toMatch(/^Classification:/m);
  });
});

describe("renderFolders", () => {
  const folder = (over: Partial<Folder> = {}): Folder => ({
    id: "f1",
    display_name: "Inbox",
    total_item_count: 100,
    unread_item_count: 5,
    ...over,
  });

  it("renders empty list", () => {
    expect(renderFolders([])).toContain("0");
  });

  it("shows display_name + counts", () => {
    const out = renderFolders([
      folder(),
      folder({ id: "f2", display_name: "Jobs", unread_item_count: 0, total_item_count: 42 }),
    ]);
    expect(out).toContain("Inbox");
    expect(out).toContain("Jobs");
    expect(out).toContain("100");
    expect(out).toContain("5");
    expect(out).toContain("42");
  });
});
