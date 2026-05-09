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
