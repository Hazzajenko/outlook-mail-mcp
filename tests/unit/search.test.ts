import { describe, expect, it } from "vitest";
import type { SearchParams } from "../../src/core/schemas.ts";
import { getEmail, listFolders, search } from "../../src/core/search.ts";
import folderFixture from "../fixtures/graph-folder.json" with { type: "json" };
import messageFixture from "../fixtures/graph-message.json" with { type: "json" };
import { FakeGraphClient } from "../helpers/fake-graph-client.ts";

const params = (over: Partial<SearchParams> = {}): SearchParams => ({ top: 50, ...over });

const msg = (id: string) => ({ ...messageFixture, id });

describe("search", () => {
  it("calls /me/messages with built query and maps lean results", async () => {
    const fake = new FakeGraphClient().enqueue({ value: [msg("a"), msg("b")] });
    const result = await search(fake, params({ top: 25 }));

    expect(fake.calls[0]?.pathOrUrl).toBe("/me/messages");
    expect(fake.calls[0]?.query?.get("$top")).toBe("25");
    expect(result.results.map((r) => r.id)).toEqual(["a", "b"]);
    expect(result.total_returned).toBe(2);
    expect(result.next_cursor).toBeUndefined();
  });

  it("auto-paginates via nextLink until top is satisfied", async () => {
    const fake = new FakeGraphClient()
      .enqueue({
        value: [msg("a"), msg("b")],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/messages?$skip=2",
      })
      .enqueue({ value: [msg("c")] });
    const result = await search(fake, params({ top: 5 }));

    expect(fake.calls).toHaveLength(2);
    expect(fake.calls[1]?.pathOrUrl).toBe("https://graph.microsoft.com/v1.0/me/messages?$skip=2");
    expect(result.results.map((r) => r.id)).toEqual(["a", "b", "c"]);
    expect(result.next_cursor).toBeUndefined();
  });

  it("returns next_cursor=nextLink when stopped at clean page boundary", async () => {
    const nextLink = "https://graph.microsoft.com/v1.0/me/messages?$skip=2&$top=2";
    const fake = new FakeGraphClient().enqueue({
      value: [msg("a"), msg("b")],
      "@odata.nextLink": nextLink,
    });
    const result = await search(fake, params({ top: 2 }));

    expect(result.results).toHaveLength(2);
    expect(result.next_cursor).toBe(nextLink);
  });

  it("next_cursor undefined when truncating mid-page (lossy)", async () => {
    const fake = new FakeGraphClient().enqueue({
      value: [msg("a"), msg("b"), msg("c"), msg("d")],
      "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/messages?$skip=4",
    });
    const result = await search(fake, params({ top: 2 }));

    expect(result.results.map((r) => r.id)).toEqual(["a", "b"]);
    expect(result.next_cursor).toBeUndefined();
  });

  it("cursor is fetched as start URL with no rebuilt query", async () => {
    const cursor = "https://graph.microsoft.com/v1.0/me/messages?$skip=50&$top=50";
    const fake = new FakeGraphClient().enqueue({ value: [msg("c")] });
    await search(fake, params({ cursor }));

    expect(fake.calls).toHaveLength(1);
    expect(fake.calls[0]?.pathOrUrl).toBe(cursor);
    expect(fake.calls[0]?.query).toBeUndefined();
  });

  it("uses well-known folder path directly (case-insensitive)", async () => {
    const fake = new FakeGraphClient().enqueue({ value: [] });
    await search(fake, params({ folder: "Inbox" }));

    expect(fake.calls[0]?.pathOrUrl).toBe("/me/mailFolders/inbox/messages");
    // no folder list call
    expect(fake.calls).toHaveLength(1);
  });

  it("resolves custom folder name via listFolders", async () => {
    const fake = new FakeGraphClient()
      .enqueue({ value: [{ ...folderFixture, id: "fold-jobs", displayName: "Jobs" }] })
      .enqueue({ value: [msg("a")] });
    const result = await search(fake, params({ folder: "Jobs" }));

    expect(fake.calls[0]?.pathOrUrl).toBe("/me/mailFolders");
    expect(fake.calls[1]?.pathOrUrl).toBe("/me/mailFolders/fold-jobs/messages");
    expect(result.results).toHaveLength(1);
  });

  it("custom folder match is case-insensitive on displayName", async () => {
    const fake = new FakeGraphClient()
      .enqueue({ value: [{ ...folderFixture, id: "fold-jobs", displayName: "Jobs" }] })
      .enqueue({ value: [] });
    await search(fake, params({ folder: "jobs" }));

    expect(fake.calls[1]?.pathOrUrl).toBe("/me/mailFolders/fold-jobs/messages");
  });

  it("throws when custom folder name not found", async () => {
    const fake = new FakeGraphClient().enqueue({
      value: [{ ...folderFixture, id: "fold-x", displayName: "Something Else" }],
    });
    await expect(search(fake, params({ folder: "DoesNotExist" }))).rejects.toThrow(
      /folder not found/i,
    );
  });
});

describe("getEmail", () => {
  it("calls /me/messages/{id} and returns FullMessage", async () => {
    const fake = new FakeGraphClient().enqueue(messageFixture);
    const result = await getEmail(fake, "AAMkADYzAA");

    expect(fake.calls[0]?.pathOrUrl).toBe("/me/messages/AAMkADYzAA");
    expect(fake.calls[0]?.query?.get("$select")).toContain("body");
    expect(result.id).toBe("AAMkADYzAA");
    expect(result.body).toBe("<html><body>Dear candidate...</body></html>");
    expect(result.body_content_type).toBe("html");
  });

  it("requests internetMessageHeaders via $select", async () => {
    const fake = new FakeGraphClient().enqueue(messageFixture);
    await getEmail(fake, "AAMkADYzAA");

    expect(fake.calls[0]?.query?.get("$select")).toContain("internetMessageHeaders");
  });

  it('sends Prefer: outlook.body-content-type="text" by default', async () => {
    const fake = new FakeGraphClient().enqueue(messageFixture);
    await getEmail(fake, "AAMkADYzAA");

    expect(fake.calls[0]?.headers).toEqual({
      Prefer: 'outlook.body-content-type="text"',
    });
  });

  it("sends Prefer html when body_format='html'", async () => {
    const fake = new FakeGraphClient().enqueue(messageFixture);
    await getEmail(fake, "AAMkADYzAA", { body_format: "html" });

    expect(fake.calls[0]?.headers).toEqual({
      Prefer: 'outlook.body-content-type="html"',
    });
  });

  it("body_format='text' is explicit (same as default)", async () => {
    const fake = new FakeGraphClient().enqueue(messageFixture);
    await getEmail(fake, "AAMkADYzAA", { body_format: "text" });

    expect(fake.calls[0]?.headers).toEqual({
      Prefer: 'outlook.body-content-type="text"',
    });
  });
});

describe("listFolders", () => {
  it("calls /me/mailFolders and maps results", async () => {
    const fake = new FakeGraphClient().enqueue({ value: [folderFixture] });
    const result = await listFolders(fake);

    expect(fake.calls[0]?.pathOrUrl).toBe("/me/mailFolders");
    expect(result).toHaveLength(1);
    expect(result[0]?.display_name).toBe("Jobs");
  });

  it("paginates folders if nextLink present", async () => {
    const fake = new FakeGraphClient()
      .enqueue({
        value: [{ ...folderFixture, id: "f1", displayName: "Inbox" }],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/mailFolders?$skip=1",
      })
      .enqueue({ value: [{ ...folderFixture, id: "f2", displayName: "Sent" }] });
    const result = await listFolders(fake);

    expect(result.map((f) => f.display_name)).toEqual(["Inbox", "Sent"]);
  });
});
