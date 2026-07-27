import { describe, expect, it } from "vitest";
import type { ParsedDateInput } from "../../src/core/date-input.ts";
import { GraphHttpError } from "../../src/core/graph-client.ts";
import type {
  CountParams,
  FilterParams,
  ListBriefParams,
  SearchParams,
} from "../../src/core/schemas.ts";
import {
  countEmails,
  getConversation,
  getEmail,
  listEmailsBrief,
  listFolders,
  search,
} from "../../src/core/search.ts";
import folderFixture from "../fixtures/graph-folder.json" with { type: "json" };
import messageFixture from "../fixtures/graph-message.json" with { type: "json" };
import { FakeGraphClient } from "../helpers/fake-graph-client.ts";

const params = (over: Partial<SearchParams> = {}): SearchParams => ({ top: 50, ...over });

const countP = (over: Partial<CountParams> = {}): CountParams => ({ ...over }) as FilterParams;

const briefP = (over: Partial<ListBriefParams> = {}): ListBriefParams => ({ top: 500, ...over });

const msg = (id: string) => ({ ...messageFixture, id });

describe("search", () => {
  it("calls /me/messages with built query and maps lean results", async () => {
    const fake = new FakeGraphClient().enqueue({ value: [msg("a"), msg("b")] });
    const result = await search(fake, params({ top: 25 }));

    expect(fake.calls[0]?.pathOrUrl).toBe("/me/messages");
    expect(fake.calls[0]?.query?.get("$top")).toBe("25");
    expect(result.results.map((r) => r.id)).toEqual(["a", "b"]);
    expect(result.total_returned).toBe(2);
    expect(result.has_more).toBe(false);
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
    expect(result.has_more).toBe(false);
    expect(result.next_cursor).toBeUndefined();
  });

  it("returns next_cursor=nextLink and has_more=true when stopped at clean page boundary", async () => {
    const nextLink = "https://graph.microsoft.com/v1.0/me/messages?$skip=2&$top=2";
    const fake = new FakeGraphClient().enqueue({
      value: [msg("a"), msg("b")],
      "@odata.nextLink": nextLink,
    });
    const result = await search(fake, params({ top: 2 }));

    expect(result.results).toHaveLength(2);
    expect(result.has_more).toBe(true);
    expect(result.next_cursor).toBe(nextLink);
  });

  it("has_more=true but next_cursor undefined when truncating mid-page (lossy)", async () => {
    const fake = new FakeGraphClient().enqueue({
      value: [msg("a"), msg("b"), msg("c"), msg("d")],
      "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/messages?$skip=4",
    });
    const result = await search(fake, params({ top: 2 }));

    expect(result.results.map((r) => r.id)).toEqual(["a", "b"]);
    expect(result.has_more).toBe(true);
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

  describe("folder name resolution", () => {
    it("populates folder display name from parentFolderId via listFolders", async () => {
      const fake = new FakeGraphClient()
        .enqueue({ value: [{ ...msg("a"), parentFolderId: "fold-inbox" }] })
        .enqueue({ value: [{ ...folderFixture, id: "fold-inbox", displayName: "Inbox" }] });
      const result = await search(fake, params());

      expect(fake.calls).toHaveLength(2);
      expect(fake.calls[1]?.pathOrUrl).toBe("/me/mailFolders");
      expect(result.results[0]?.folder).toBe("Inbox");
    });

    it("skips listFolders when no result has parentFolderId", async () => {
      const fake = new FakeGraphClient().enqueue({ value: [msg("a"), msg("b")] });
      const result = await search(fake, params());

      expect(fake.calls).toHaveLength(1);
      expect(result.results[0]).not.toHaveProperty("folder");
    });

    it("leaves folder undefined when listFolders doesn't return the id", async () => {
      const fake = new FakeGraphClient()
        .enqueue({ value: [{ ...msg("a"), parentFolderId: "fold-unknown" }] })
        .enqueue({ value: [{ ...folderFixture, id: "fold-other", displayName: "Other" }] });
      const result = await search(fake, params());

      expect(result.results[0]).not.toHaveProperty("folder");
    });

    it("strips noisy zero-width chars from body_preview", async () => {
      const fake = new FakeGraphClient().enqueue({
        value: [{ ...msg("a"), bodyPreview: "Hello​‌‍world͏!" }],
      });
      const result = await search(fake, params());

      expect(result.results[0]?.body_preview).toBe("Helloworld!");
    });
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

  describe("header trimming", () => {
    const noisy = {
      ...(messageFixture as Record<string, unknown>),
      internetMessageHeaders: [
        { name: "Authentication-Results", value: "spf=pass" },
        { name: "Return-Path", value: "<bounces@example.com>" },
        { name: "reply-to", value: "<human@example.com>" },
        { name: "List-Unsubscribe", value: "<https://example.com/u>" },
        { name: "X-Microsoft-Antispam-Message-Info", value: "x".repeat(2000) },
        { name: "DKIM-Signature", value: "v=1; a=rsa-sha256; ..." },
        { name: "X-MS-Exchange-Organization-SCL", value: "1" },
      ],
    };

    it("keeps only the notable headers by default", async () => {
      const fake = new FakeGraphClient().enqueue(noisy);
      const result = await getEmail(fake, "AAMkADYzAA");

      expect(result.internet_message_headers.map((h) => h.name)).toEqual([
        "Authentication-Results",
        "Return-Path",
        "reply-to",
        "List-Unsubscribe",
      ]);
    });

    it("matches header names case-insensitively (senders vary the casing)", async () => {
      const fake = new FakeGraphClient().enqueue(noisy);
      const result = await getEmail(fake, "AAMkADYzAA");

      expect(result.internet_message_headers).toContainEqual({
        name: "reply-to",
        value: "<human@example.com>",
      });
    });

    it("returns every header when include_all_headers is set", async () => {
      const fake = new FakeGraphClient().enqueue(noisy);
      const result = await getEmail(fake, "AAMkADYzAA", { include_all_headers: true });

      expect(result.internet_message_headers).toHaveLength(7);
    });

    it("still requests them via $select — Graph cannot select a subset", async () => {
      const fake = new FakeGraphClient().enqueue(noisy);
      await getEmail(fake, "AAMkADYzAA");

      expect(fake.calls[0]?.query?.get("$select")).toContain("internetMessageHeaders");
    });

    it("leaves the two headers the CLI renders intact", async () => {
      const fake = new FakeGraphClient().enqueue(noisy);
      const result = await getEmail(fake, "AAMkADYzAA");

      const names = result.internet_message_headers.map((h) => h.name);
      expect(names).toContain("Authentication-Results");
      expect(names).toContain("Return-Path");
    });
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

describe("getConversation", () => {
  const convoMsg = (id: string, receivedAt: string) => ({
    ...messageFixture,
    id,
    receivedDateTime: receivedAt,
    conversationId: "conv-x",
  });

  it("filters by conversationId and returns FullMessage array", async () => {
    const fake = new FakeGraphClient().enqueue({
      value: [convoMsg("a", "2026-05-08T10:00:00Z"), convoMsg("b", "2026-05-08T11:00:00Z")],
    });
    const result = await getConversation(fake, "conv-x");

    expect(fake.calls[0]?.pathOrUrl).toBe("/me/messages");
    expect(fake.calls[0]?.query?.get("$filter")).toBe("conversationId eq 'conv-x'");
    expect(result.map((m) => m.id)).toEqual(["a", "b"]);
    expect(result[0]?.body).toBeDefined();
  });

  it("sorts results ascending by received_at (chronological)", async () => {
    const fake = new FakeGraphClient().enqueue({
      value: [
        convoMsg("c", "2026-05-08T12:00:00Z"),
        convoMsg("a", "2026-05-08T10:00:00Z"),
        convoMsg("b", "2026-05-08T11:00:00Z"),
      ],
    });
    const result = await getConversation(fake, "conv-x");

    expect(result.map((m) => m.id)).toEqual(["a", "b", "c"]);
  });

  it("does not send $orderby (would trigger InefficientFilter on Graph)", async () => {
    const fake = new FakeGraphClient().enqueue({ value: [] });
    await getConversation(fake, "conv-x");

    expect(fake.calls[0]?.query?.has("$orderby")).toBe(false);
  });

  it('sends Prefer: outlook.body-content-type="text" by default', async () => {
    const fake = new FakeGraphClient().enqueue({ value: [] });
    await getConversation(fake, "conv-x");

    expect(fake.calls[0]?.headers).toEqual({
      Prefer: 'outlook.body-content-type="text"',
    });
  });

  it("sends html Prefer when body_format='html'", async () => {
    const fake = new FakeGraphClient().enqueue({ value: [] });
    await getConversation(fake, "conv-x", { body_format: "html" });

    expect(fake.calls[0]?.headers).toEqual({
      Prefer: 'outlook.body-content-type="html"',
    });
  });

  it("escapes single quotes in conversationId per OData", async () => {
    const fake = new FakeGraphClient().enqueue({ value: [] });
    await getConversation(fake, "weird'id");

    expect(fake.calls[0]?.query?.get("$filter")).toBe("conversationId eq 'weird''id'");
  });

  it("paginates via nextLink", async () => {
    const fake = new FakeGraphClient()
      .enqueue({
        value: [convoMsg("a", "2026-05-08T10:00:00Z")],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/messages?$skip=1",
      })
      .enqueue({ value: [convoMsg("b", "2026-05-08T11:00:00Z")] });
    const result = await getConversation(fake, "conv-x");

    expect(fake.calls).toHaveLength(2);
    expect(result.map((m) => m.id)).toEqual(["a", "b"]);
  });

  it("re-sends Prefer header on paginated nextLink calls", async () => {
    const fake = new FakeGraphClient()
      .enqueue({
        value: [convoMsg("a", "2026-05-08T10:00:00Z")],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/messages?$skip=1",
      })
      .enqueue({ value: [convoMsg("b", "2026-05-08T11:00:00Z")] });
    await getConversation(fake, "conv-x");

    expect(fake.calls[1]?.headers).toEqual({
      Prefer: 'outlook.body-content-type="text"',
    });
  });

  it("respects top cap and stops paginating", async () => {
    const fake = new FakeGraphClient().enqueue({
      value: [
        convoMsg("a", "2026-05-08T10:00:00Z"),
        convoMsg("b", "2026-05-08T11:00:00Z"),
        convoMsg("c", "2026-05-08T12:00:00Z"),
      ],
      "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/messages?$skip=3",
    });
    const result = await getConversation(fake, "conv-x", { top: 2 });

    expect(fake.calls).toHaveLength(1);
    expect(result).toHaveLength(2);
  });
});

describe("countEmails", () => {
  it("returns @odata.count from Graph", async () => {
    const fake = new FakeGraphClient().enqueue({ "@odata.count": 1247, value: [] });
    const result = await countEmails(fake, countP({ is_unread: true }));

    expect(result.count).toBe(1247);
  });

  it("sends ConsistencyLevel: eventual header", async () => {
    const fake = new FakeGraphClient().enqueue({ "@odata.count": 0, value: [] });
    await countEmails(fake, countP());

    expect(fake.calls[0]?.headers).toEqual({ ConsistencyLevel: "eventual" });
  });

  it("uses folder-scoped endpoint when folder given", async () => {
    const fake = new FakeGraphClient().enqueue({ "@odata.count": 50, value: [] });
    await countEmails(fake, countP({ folder: "Inbox" }));

    expect(fake.calls[0]?.pathOrUrl).toBe("/me/mailFolders/inbox/messages");
  });

  it("resolves custom folder name via listFolders", async () => {
    const fake = new FakeGraphClient()
      .enqueue({ value: [{ ...folderFixture, id: "fold-jobs", displayName: "Jobs" }] })
      .enqueue({ "@odata.count": 12, value: [] });
    const result = await countEmails(fake, countP({ folder: "Jobs" }));

    expect(fake.calls[0]?.pathOrUrl).toBe("/me/mailFolders");
    expect(fake.calls[1]?.pathOrUrl).toBe("/me/mailFolders/fold-jobs/messages");
    expect(result.count).toBe(12);
  });
});

describe("listEmailsBrief", () => {
  const briefMsg = (id: string, receivedAt: string, address: string, subject: string) => ({
    id,
    receivedDateTime: receivedAt,
    subject,
    from: { emailAddress: { address } },
  });

  it("returns newline-joined text with date | from | subject per line", async () => {
    const fake = new FakeGraphClient().enqueue({
      value: [
        briefMsg("a", "2026-05-08T14:43:14Z", "recruiter@goldman.com", "Assessment invitation"),
        briefMsg("b", "2026-05-07T09:00:00Z", "alerts@linkedin.com", "5 new jobs match"),
      ],
    });
    const result = await listEmailsBrief(fake, briefP());

    expect(result.lines).toBe(
      "2026-05-08 14:43 | recruiter@goldman.com | Assessment invitation\n" +
        "2026-05-07 09:00 | alerts@linkedin.com | 5 new jobs match",
    );
    expect(result.total_returned).toBe(2);
    expect(result.has_more).toBe(false);
  });

  it("truncates long subjects to 80 chars with ellipsis", async () => {
    const longSubject = "x".repeat(120);
    const fake = new FakeGraphClient().enqueue({
      value: [briefMsg("a", "2026-05-08T00:00:00Z", "x@y.com", longSubject)],
    });
    const result = await listEmailsBrief(fake, briefP());

    const subjectFromLine = result.lines.split(" | ")[2];
    expect(subjectFromLine).toHaveLength(80);
    expect(subjectFromLine?.endsWith("…")).toBe(true);
  });

  it("auto-paginates until top reached", async () => {
    const fake = new FakeGraphClient()
      .enqueue({
        value: [briefMsg("a", "2026-05-08T10:00:00Z", "a@x.com", "A")],
        "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/messages?$skip=1",
      })
      .enqueue({ value: [briefMsg("b", "2026-05-07T10:00:00Z", "b@x.com", "B")] });
    const result = await listEmailsBrief(fake, briefP({ top: 5 }));

    expect(fake.calls).toHaveLength(2);
    expect(result.total_returned).toBe(2);
    expect(result.has_more).toBe(false);
  });

  it("stops at top and reports has_more=true when more pages remain", async () => {
    const fake = new FakeGraphClient().enqueue({
      value: [
        briefMsg("a", "2026-05-08T10:00:00Z", "a@x.com", "A"),
        briefMsg("b", "2026-05-07T10:00:00Z", "b@x.com", "B"),
      ],
      "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/messages?$skip=2",
    });
    const result = await listEmailsBrief(fake, briefP({ top: 2 }));

    expect(result.total_returned).toBe(2);
    expect(result.has_more).toBe(true);
  });

  it("truncates mid-page and sets has_more=true", async () => {
    const fake = new FakeGraphClient().enqueue({
      value: [
        briefMsg("a", "2026-05-08T10:00:00Z", "a@x.com", "A"),
        briefMsg("b", "2026-05-07T10:00:00Z", "b@x.com", "B"),
        briefMsg("c", "2026-05-06T10:00:00Z", "c@x.com", "C"),
      ],
      "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/messages?$skip=3",
    });
    const result = await listEmailsBrief(fake, briefP({ top: 2 }));

    expect(result.total_returned).toBe(2);
    expect(result.has_more).toBe(true);
    expect(result.lines.split("\n")).toHaveLength(2);
  });

  it("handles missing from gracefully", async () => {
    const fake = new FakeGraphClient().enqueue({
      value: [{ id: "a", receivedDateTime: "2026-05-08T10:00:00Z", subject: "x", from: null }],
    });
    const result = await listEmailsBrief(fake, briefP());

    expect(result.lines).toContain("(unknown)");
  });

  it("collapses whitespace in subject", async () => {
    const fake = new FakeGraphClient().enqueue({
      value: [briefMsg("a", "2026-05-08T10:00:00Z", "x@y.com", "  hello\n\t world  ")],
    });
    const result = await listEmailsBrief(fake, briefP());

    expect(result.lines).toContain("| hello world");
  });
});

describe("InefficientFilter translation", () => {
  const inefficient = () =>
    new GraphHttpError(
      400,
      "Bad Request",
      '{"error":{"code":"InefficientFilter","message":"The restriction or sort order is too complex for this operation."}}',
    );

  const bound = (): ParsedDateInput => ({
    date: new Date("2026-05-01T00:00:00Z"),
    dateOnly: false,
  });

  it("explains the missing date bound on search", async () => {
    const fake = new FakeGraphClient().enqueueError(inefficient());

    await expect(search(fake, params({ inference_classification: "focused" }))).rejects.toThrow(
      /Pass since and\/or until/,
    );
  });

  it("explains the missing date bound on listEmailsBrief", async () => {
    const fake = new FakeGraphClient().enqueueError(inefficient());

    await expect(
      listEmailsBrief(fake, briefP({ inference_classification: "other" })),
    ).rejects.toThrow(/Pass since and\/or until/);
  });

  it("names the offending value and keeps the original error as cause", async () => {
    const original = inefficient();
    const fake = new FakeGraphClient().enqueueError(original);

    const err = await search(fake, params({ inference_classification: "focused" })).catch(
      (e: unknown) => e,
    );

    expect((err as Error).message).toContain("inference_classification='focused'");
    expect((err as Error).cause).toBe(original);
  });

  // The hint would be wrong here — a bound is already present, so the 400 has
  // some other cause and must not be papered over with misleading advice.
  it("rethrows untouched when a date bound is already present", async () => {
    const original = inefficient();
    const fake = new FakeGraphClient().enqueueError(original);

    const err = await search(
      fake,
      params({ inference_classification: "focused", since: bound() }),
    ).catch((e: unknown) => e);

    expect(err).toBe(original);
  });

  it("rethrows untouched when inference_classification is absent", async () => {
    const original = inefficient();
    const fake = new FakeGraphClient().enqueueError(original);

    const err = await search(fake, params()).catch((e: unknown) => e);

    expect(err).toBe(original);
  });

  it("rethrows unrelated Graph errors untouched", async () => {
    const original = new GraphHttpError(429, "Too Many Requests", '{"error":{"code":"Throttled"}}');
    const fake = new FakeGraphClient().enqueueError(original);

    const err = await search(fake, params({ inference_classification: "focused" })).catch(
      (e: unknown) => e,
    );

    expect(err).toBe(original);
  });

  it("leaves countEmails unwrapped — it does not sort, so it needs no bound", async () => {
    const fake = new FakeGraphClient().enqueue({ "@odata.count": 7 });

    const result = await countEmails(fake, countP({ inference_classification: "focused" }));

    expect(result.count).toBe(7);
  });
});
