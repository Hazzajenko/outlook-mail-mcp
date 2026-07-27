export interface GraphClient {
  get(
    pathOrUrl: string,
    query?: URLSearchParams,
    headers?: Record<string, string>,
  ): Promise<unknown>;
}

export interface GraphPage {
  value: unknown[];
  "@odata.nextLink"?: string;
}

/**
 * Non-2xx response from Graph. Lives with the protocol rather than the HTTP
 * impl so core logic can catch it without importing HttpGraphClient — any
 * GraphClient implementation is expected to signal failures with this.
 */
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
