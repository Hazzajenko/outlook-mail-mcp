import { describe, expect, it } from "vitest";
import { FakeGraphClient } from "../helpers/fake-graph-client.ts";

describe("FakeGraphClient", () => {
  it("returns enqueued responses in order", async () => {
    const fake = new FakeGraphClient().enqueue({ a: 1 }).enqueue({ b: 2 });
    expect(await fake.get("/x")).toEqual({ a: 1 });
    expect(await fake.get("/y")).toEqual({ b: 2 });
  });

  it("records calls with path and query", async () => {
    const fake = new FakeGraphClient().enqueue({});
    const q = new URLSearchParams({ $top: "10" });
    await fake.get("/me/messages", q);
    expect(fake.calls).toHaveLength(1);
    expect(fake.calls[0]?.pathOrUrl).toBe("/me/messages");
    expect(fake.calls[0]?.query?.get("$top")).toBe("10");
  });

  it("throws on call when queue is empty", async () => {
    const fake = new FakeGraphClient();
    await expect(fake.get("/anything")).rejects.toThrow(/unexpected call/);
  });
});
