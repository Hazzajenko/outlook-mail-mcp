import { describe, expect, it } from "vitest";
import { mergeEnvContent } from "../../src/core/env-file.ts";

describe("mergeEnvContent", () => {
  it("creates content from scratch when no existing file", () => {
    expect(mergeEnvContent(undefined, { OUTLOOK_QUERY_CLIENT_ID: "abc" })).toBe(
      "OUTLOOK_QUERY_CLIENT_ID=abc\n",
    );
  });

  it("appends missing keys to existing content", () => {
    const existing = "OTHER=1\n";
    expect(mergeEnvContent(existing, { OUTLOOK_QUERY_CLIENT_ID: "abc" })).toBe(
      "OTHER=1\nOUTLOOK_QUERY_CLIENT_ID=abc\n",
    );
  });

  it("replaces an existing key in place", () => {
    const existing = "A=1\nOUTLOOK_QUERY_CLIENT_ID=old\nB=2\n";
    expect(mergeEnvContent(existing, { OUTLOOK_QUERY_CLIENT_ID: "new" })).toBe(
      "A=1\nOUTLOOK_QUERY_CLIENT_ID=new\nB=2\n",
    );
  });

  it("preserves comments and blank lines", () => {
    const existing = "# my config\n\nOTHER=1\n";
    expect(mergeEnvContent(existing, { OUTLOOK_QUERY_CLIENT_ID: "abc" })).toBe(
      "# my config\n\nOTHER=1\nOUTLOOK_QUERY_CLIENT_ID=abc\n",
    );
  });

  it("removes duplicate occurrences of a replaced key", () => {
    const existing = "OUTLOOK_QUERY_CLIENT_ID=one\nOUTLOOK_QUERY_CLIENT_ID=two\n";
    expect(mergeEnvContent(existing, { OUTLOOK_QUERY_CLIENT_ID: "new" })).toBe(
      "OUTLOOK_QUERY_CLIENT_ID=new\n",
    );
  });

  it("handles missing trailing newline in existing content", () => {
    expect(mergeEnvContent("OTHER=1", { K: "v" })).toBe("OTHER=1\nK=v\n");
  });

  it("writes multiple keys in given order", () => {
    expect(
      mergeEnvContent(undefined, {
        OUTLOOK_QUERY_CLIENT_ID: "abc",
        OUTLOOK_QUERY_TENANT_ID: "consumers",
      }),
    ).toBe("OUTLOOK_QUERY_CLIENT_ID=abc\nOUTLOOK_QUERY_TENANT_ID=consumers\n");
  });

  it("deletes a key when its value is null", () => {
    const existing = "A=1\nOUTLOOK_QUERY_TENANT_ID=consumers\nB=2\n";
    expect(mergeEnvContent(existing, { OUTLOOK_QUERY_TENANT_ID: null })).toBe("A=1\nB=2\n");
  });

  it("null for an absent key writes nothing", () => {
    expect(mergeEnvContent("A=1\n", { K: null })).toBe("A=1\n");
  });

  it("matches keys with surrounding whitespace around =", () => {
    const existing = "OUTLOOK_QUERY_CLIENT_ID = old\n";
    expect(mergeEnvContent(existing, { OUTLOOK_QUERY_CLIENT_ID: "new" })).toBe(
      "OUTLOOK_QUERY_CLIENT_ID=new\n",
    );
  });
});
