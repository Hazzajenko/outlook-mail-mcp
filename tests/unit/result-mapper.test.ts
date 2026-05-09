import { describe, expect, it } from "vitest";
import { mapFolder, mapFullMessage, mapLeanMessage } from "../../src/core/result-mapper.ts";
import folderFixture from "../fixtures/graph-folder.json" with { type: "json" };
import messageFixture from "../fixtures/graph-message.json" with { type: "json" };

describe("mapLeanMessage", () => {
  it("maps a Graph message to LeanMessage shape", () => {
    const m = mapLeanMessage(messageFixture);
    expect(m).toEqual({
      id: "AAMkADYzAA",
      from: { name: "Example Recruiter", address: "recruiter@example.com" },
      subject: "Assessment invitation",
      received_at: "2026-05-08T14:43:14Z",
      has_attachment: false,
      is_read: false,
      conversation_id: "AAQkADYz",
      web_link: "https://outlook.office.com/mail/inbox/id/AAMkADYzAA",
      body_preview: "Dear candidate, please complete the assessment by Friday.",
    });
  });

  it("falls back to empty address when from missing", () => {
    const m = mapLeanMessage({ ...messageFixture, from: null });
    expect(m.from).toEqual({ address: "" });
  });

  it("falls back to empty subject when null", () => {
    const m = mapLeanMessage({ ...messageFixture, subject: null });
    expect(m.subject).toBe("");
  });

  it("falls back to empty body_preview when null", () => {
    const m = mapLeanMessage({ ...messageFixture, bodyPreview: null });
    expect(m.body_preview).toBe("");
  });
});

describe("mapFullMessage", () => {
  it("maps a Graph message to FullMessage shape", () => {
    const m = mapFullMessage(messageFixture);
    expect(m.id).toBe("AAMkADYzAA");
    expect(m.body).toBe("<html><body>Dear candidate...</body></html>");
    expect(m.body_content_type).toBe("html");
    expect(m.importance).toBe("high");
    expect(m.to).toEqual([{ name: "Test Recipient", address: "recipient@example.com" }]);
    expect(m.cc).toEqual([{ address: "team@example.com" }]);
  });

  it("lowercases contentType (HTML -> html, Text -> text)", () => {
    const text = mapFullMessage({
      ...messageFixture,
      body: { contentType: "Text", content: "plain text body" },
    });
    expect(text.body_content_type).toBe("text");
  });

  it("defaults importance to 'normal' when missing", () => {
    const m = mapFullMessage({ ...messageFixture, importance: undefined });
    expect(m.importance).toBe("normal");
  });

  it("empty arrays for missing toRecipients/ccRecipients", () => {
    const m = mapFullMessage({
      ...messageFixture,
      toRecipients: undefined,
      ccRecipients: undefined,
    });
    expect(m.to).toEqual([]);
    expect(m.cc).toEqual([]);
  });

  it("defaults body to empty string when missing", () => {
    const m = mapFullMessage({ ...messageFixture, body: undefined });
    expect(m.body).toBe("");
    expect(m.body_content_type).toBe("text");
  });
});

describe("mapFolder", () => {
  it("maps a Graph mailFolder to Folder shape", () => {
    expect(mapFolder(folderFixture)).toEqual({
      id: "AAMkAD-folder",
      display_name: "Jobs",
      parent_folder_id: "AAMk-parent",
      total_item_count: 42,
      unread_item_count: 5,
    });
  });

  it("omits parent_folder_id when null/missing", () => {
    const f = mapFolder({ ...folderFixture, parentFolderId: null });
    expect(f.parent_folder_id).toBeUndefined();
  });
});
