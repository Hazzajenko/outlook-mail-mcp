import type { GraphClient } from "./graph-client.ts";

const DEFAULT_BASE_URL = "https://graph.microsoft.com/v1.0";

export interface HttpGraphClientOptions {
  getToken: () => Promise<string>;
  baseUrl?: string;
  fetch?: typeof fetch;
}

export class HttpGraphClient implements GraphClient {
  private readonly getToken: () => Promise<string>;
  private readonly baseUrl: string;
  private readonly fetchFn: typeof fetch;

  constructor(opts: HttpGraphClientOptions) {
    this.getToken = opts.getToken;
    this.baseUrl = opts.baseUrl ?? DEFAULT_BASE_URL;
    this.fetchFn = opts.fetch ?? globalThis.fetch.bind(globalThis);
  }

  async get(pathOrUrl: string, query?: URLSearchParams): Promise<unknown> {
    const base = pathOrUrl.startsWith("http") ? pathOrUrl : `${this.baseUrl}${pathOrUrl}`;
    const url = query ? `${base}?${query.toString()}` : base;

    const token = await this.getToken();
    const res = await this.fetchFn(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/json",
      },
    });

    if (!res.ok) {
      const body = await res.text().catch(() => "");
      throw new GraphHttpError(res.status, res.statusText, body);
    }

    return res.json();
  }
}

export class GraphHttpError extends Error {
  readonly status: number;
  readonly statusText: string;
  readonly body: string;

  constructor(status: number, statusText: string, body: string) {
    super(`Graph ${status} ${statusText}: ${body.slice(0, 500)}`);
    this.name = "GraphHttpError";
    this.status = status;
    this.statusText = statusText;
    this.body = body;
  }
}
