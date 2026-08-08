import { type GraphClient, GraphHttpError } from "./graph-client.ts";

// Re-exported for callers that already import it from here.
export { GraphHttpError };

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

  async get(
    pathOrUrl: string,
    query?: URLSearchParams,
    headers?: Record<string, string>,
  ): Promise<unknown> {
    const base = pathOrUrl.startsWith("http") ? pathOrUrl : `${this.baseUrl}${pathOrUrl}`;
    const url = query ? `${base}?${query.toString()}` : base;

    const token = await this.getToken();
    const res = await this.fetchFn(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/json",
        ...(headers ?? {}),
      },
    });

    if (!res.ok) {
      const body = await res.text().catch(() => "");
      throw new GraphHttpError(res.status, res.statusText, body);
    }

    return res.json();
  }
}
