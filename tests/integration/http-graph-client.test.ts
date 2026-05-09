import { HttpResponse, http } from "msw";
import { setupServer } from "msw/node";
import { afterAll, afterEach, beforeAll, describe, expect, it } from "vitest";
import { GraphHttpError, HttpGraphClient } from "../../src/core/http-graph-client.ts";

const server = setupServer();

beforeAll(() => server.listen({ onUnhandledRequest: "error" }));
afterEach(() => server.resetHandlers());
afterAll(() => server.close());

const makeClient = (token = "test-token") => new HttpGraphClient({ getToken: async () => token });

describe("HttpGraphClient", () => {
  describe("URL construction", () => {
    it("relative path -> base + path", async () => {
      let capturedPath = "";
      server.use(
        http.get("https://graph.microsoft.com/v1.0/me/messages", ({ request }) => {
          capturedPath = new URL(request.url).pathname;
          return HttpResponse.json({ value: [] });
        }),
      );

      await makeClient().get("/me/messages");
      expect(capturedPath).toBe("/v1.0/me/messages");
    });

    it("absolute URL passed through unchanged", async () => {
      server.use(
        http.get("https://graph.microsoft.com/v1.0/me/messages", ({ request }) => {
          const u = new URL(request.url);
          return HttpResponse.json({ value: [], skip: u.searchParams.get("$skip") });
        }),
      );

      const result = (await makeClient().get(
        "https://graph.microsoft.com/v1.0/me/messages?$skip=10",
      )) as { skip: string };
      expect(result.skip).toBe("10");
    });

    it("appends query string from URLSearchParams", async () => {
      let capturedQuery: URLSearchParams | undefined;
      server.use(
        http.get("https://graph.microsoft.com/v1.0/me/messages", ({ request }) => {
          capturedQuery = new URL(request.url).searchParams;
          return HttpResponse.json({ value: [] });
        }),
      );

      const q = new URLSearchParams();
      q.set("$top", "25");
      q.set("$select", "id,subject");
      q.set("$search", '"interview"');
      await makeClient().get("/me/messages", q);

      expect(capturedQuery?.get("$top")).toBe("25");
      expect(capturedQuery?.get("$select")).toBe("id,subject");
      expect(capturedQuery?.get("$search")).toBe('"interview"');
    });
  });

  describe("auth", () => {
    it("attaches Authorization: Bearer header", async () => {
      let capturedAuth: string | null = null;
      server.use(
        http.get("https://graph.microsoft.com/v1.0/me/messages", ({ request }) => {
          capturedAuth = request.headers.get("authorization");
          return HttpResponse.json({ value: [] });
        }),
      );

      await makeClient("my-secret").get("/me/messages");
      expect(capturedAuth).toBe("Bearer my-secret");
    });

    it("calls getToken on each request (no caching at this layer)", async () => {
      let calls = 0;
      const client = new HttpGraphClient({
        getToken: async () => {
          calls++;
          return `tok-${calls}`;
        },
      });

      const seen: string[] = [];
      server.use(
        http.get("https://graph.microsoft.com/v1.0/me/messages", ({ request }) => {
          const auth = request.headers.get("authorization") ?? "";
          seen.push(auth);
          return HttpResponse.json({ value: [] });
        }),
      );

      await client.get("/me/messages");
      await client.get("/me/messages");
      expect(seen).toEqual(["Bearer tok-1", "Bearer tok-2"]);
    });
  });

  describe("response handling", () => {
    it("returns parsed JSON on 2xx", async () => {
      server.use(
        http.get("https://graph.microsoft.com/v1.0/me/messages", () =>
          HttpResponse.json({ value: [{ id: "x" }] }),
        ),
      );
      const result = await makeClient().get("/me/messages");
      expect(result).toEqual({ value: [{ id: "x" }] });
    });

    it("throws GraphHttpError on 4xx with status", async () => {
      server.use(
        http.get("https://graph.microsoft.com/v1.0/me/messages", () =>
          HttpResponse.json(
            { error: { code: "InvalidAuth", message: "Token bad" } },
            { status: 401 },
          ),
        ),
      );

      await expect(makeClient().get("/me/messages")).rejects.toMatchObject({
        name: "GraphHttpError",
        status: 401,
      });
    });

    it("throws GraphHttpError on 5xx", async () => {
      server.use(
        http.get("https://graph.microsoft.com/v1.0/me/messages", () =>
          HttpResponse.text("upstream blew up", { status: 503 }),
        ),
      );

      const err = await makeClient()
        .get("/me/messages")
        .catch((e: unknown) => e);
      expect(err).toBeInstanceOf(GraphHttpError);
      expect((err as GraphHttpError).status).toBe(503);
    });
  });
});
