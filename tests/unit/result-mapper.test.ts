import { describe, expect, it } from "vitest";
import { mapFolder, mapFullMessage, mapLeanMessage } from "../../src/core/result-mapper.ts";
import folderFixture from "../fixtures/graph-folder.json" with { type: "json" };
import messageFixture from "../fixtures/graph-message.json" with { type: "json" };

describe("mapLeanMessage", () => {
  it("maps a Graph message to LeanMessage shape", () => {
    const m = mapLeanMessage(messageFixture);
    expect(m).toEqual({
      id: "AAMkADYzAA",
      from: { name: "Goldman Recruiter", address: "recruiter@goldman.com" },
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

  it("passes through inference_classification when present", () => {
    const m = mapLeanMessage({ ...messageFixture, inferenceClassification: "other" });
    expect(m.inference_classification).toBe("other");
  });

  it("passes through parent_folder_id when present", () => {
    const m = mapLeanMessage({ ...messageFixture, parentFolderId: "fold-junk" });
    expect(m.parent_folder_id).toBe("fold-junk");
  });

  it("omits inference_classification + parent_folder_id when missing", () => {
    const m = mapLeanMessage(messageFixture);
    expect(m).not.toHaveProperty("inference_classification");
    expect(m).not.toHaveProperty("parent_folder_id");
  });
});

describe("mapFullMessage", () => {
  it("maps a Graph message to FullMessage shape", () => {
    const m = mapFullMessage(messageFixture);
    expect(m.id).toBe("AAMkADYzAA");
    expect(m.body).toBe("<html><body>Dear candidate...</body></html>");
    expect(m.body_content_type).toBe("html");
    expect(m.importance).toBe("high");
    expect(m.to).toEqual([{ name: "Harry Jenkins", address: "jenkinsh1@outlook.com" }]);
    expect(m.cc).toEqual([{ address: "team@goldman.com" }]);
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

  it("passes through internet_message_headers when present", () => {
    const m = mapFullMessage({
      ...messageFixture,
      internetMessageHeaders: [
        { name: "Authentication-Results", value: "spf=fail; dkim=none; dmarc=fail" },
        { name: "Return-Path", value: "<bounces@soulmate4u.com>" },
      ],
    });
    expect(m.internet_message_headers).toEqual([
      { name: "Authentication-Results", value: "spf=fail; dkim=none; dmarc=fail" },
      { name: "Return-Path", value: "<bounces@soulmate4u.com>" },
    ]);
  });

  it("defaults internet_message_headers to [] when missing", () => {
    const m = mapFullMessage(messageFixture);
    expect(m.internet_message_headers).toEqual([]);
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
