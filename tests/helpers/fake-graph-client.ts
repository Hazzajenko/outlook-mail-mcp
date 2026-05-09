import type { GraphClient } from "../../src/core/graph-client.ts";

export interface FakeCall {
  pathOrUrl: string;
  query: URLSearchParams | undefined;
  headers: Record<string, string> | undefined;
}

export class FakeGraphClient implements GraphClient {
  readonly calls: FakeCall[] = [];
  private readonly responses: unknown[] = [];

  enqueue(response: unknown): this {
    this.responses.push(response);
    return this;
  }

  async get(
    pathOrUrl: string,
    query?: URLSearchParams,
    headers?: Record<string, string>,
  ): Promise<unknown> {
    this.calls.push({ pathOrUrl, query, headers });
    if (this.responses.length === 0) {
      throw new Error(`FakeGraphClient: unexpected call to ${pathOrUrl}`);
    }
    return this.responses.shift();
  }
}
