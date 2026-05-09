import type { GraphClient } from "../../src/core/graph-client.ts";

export interface FakeCall {
  pathOrUrl: string;
  query: URLSearchParams | undefined;
}

export class FakeGraphClient implements GraphClient {
  readonly calls: FakeCall[] = [];
  private readonly responses: unknown[] = [];

  enqueue(response: unknown): this {
    this.responses.push(response);
    return this;
  }

  async get(pathOrUrl: string, query?: URLSearchParams): Promise<unknown> {
    this.calls.push({ pathOrUrl, query });
    if (this.responses.length === 0) {
      throw new Error(`FakeGraphClient: unexpected call to ${pathOrUrl}`);
    }
    return this.responses.shift();
  }
}
