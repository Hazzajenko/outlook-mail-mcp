export interface GraphClient {
  get(pathOrUrl: string, query?: URLSearchParams): Promise<unknown>;
}

export interface GraphPage {
  value: unknown[];
  "@odata.nextLink"?: string;
}
