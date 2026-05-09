import { z } from "zod";
import type { GraphClient } from "./graph-client.ts";
import { buildGraphQuery, FULL_SELECT } from "./query-builder.ts";
import { mapFolder, mapFullMessage, mapLeanMessage } from "./result-mapper.ts";
import type { Folder, FullMessage, SearchParams, SearchResult } from "./schemas.ts";

const PageSchema = z.object({
  value: z.array(z.unknown()),
  "@odata.nextLink": z.string().optional(),
});

const WELL_KNOWN_FOLDERS = new Set([
  "archive",
  "clutter",
  "conflicts",
  "conversationhistory",
  "deleteditems",
  "drafts",
  "inbox",
  "junkemail",
  "localfailures",
  "msgfolderroot",
  "outbox",
  "recoverableitemsdeletions",
  "scheduled",
  "searchfolders",
  "sentitems",
  "serverfailures",
  "syncissues",
]);

export async function search(client: GraphClient, params: SearchParams): Promise<SearchResult> {
  let pagePath: string;
  let pageQuery: URLSearchParams | undefined;
  if (params.cursor) {
    pagePath = params.cursor;
    pageQuery = undefined;
  } else {
    const folderId = params.folder ? await resolveFolderId(client, params.folder) : undefined;
    const built = buildGraphQuery(params, folderId);
    pagePath = built.endpoint;
    pageQuery = built.query;
  }

  const out = [];
  let nextCursor: string | undefined;

  while (true) {
    const raw = await client.get(pagePath, pageQuery);
    const { value, "@odata.nextLink": nextLink } = PageSchema.parse(raw);
    const remaining = params.top - out.length;

    if (value.length > remaining) {
      // truncating mid-page: nextLink would skip unconsumed items, so don't expose a cursor
      for (let i = 0; i < remaining; i++) {
        out.push(mapLeanMessage(value[i]));
      }
      break;
    }

    for (const v of value) out.push(mapLeanMessage(v));

    if (out.length >= params.top) {
      nextCursor = nextLink;
      break;
    }
    if (nextLink === undefined) break;

    pagePath = nextLink;
    pageQuery = undefined;
  }

  return {
    results: out,
    total_returned: out.length,
    ...(nextCursor !== undefined ? { next_cursor: nextCursor } : {}),
  };
}

export interface GetEmailOptions {
  body_format?: "text" | "html";
}

export async function getEmail(
  client: GraphClient,
  id: string,
  opts: GetEmailOptions = {},
): Promise<FullMessage> {
  const query = new URLSearchParams();
  query.set("$select", FULL_SELECT);
  const format = opts.body_format ?? "text";
  const headers = { Prefer: `outlook.body-content-type="${format}"` };
  const raw = await client.get(`/me/messages/${id}`, query, headers);
  return mapFullMessage(raw);
}

export async function listFolders(client: GraphClient): Promise<Folder[]> {
  const out: Folder[] = [];
  let pagePath: string = "/me/mailFolders";
  let pageQuery: URLSearchParams | undefined;

  while (true) {
    const raw = await client.get(pagePath, pageQuery);
    const { value, "@odata.nextLink": nextLink } = PageSchema.parse(raw);
    for (const v of value) out.push(mapFolder(v));
    if (nextLink === undefined) break;
    pagePath = nextLink;
    pageQuery = undefined;
  }

  return out;
}

async function resolveFolderId(client: GraphClient, name: string): Promise<string> {
  const lc = name.toLowerCase();
  if (WELL_KNOWN_FOLDERS.has(lc)) return lc;
  const folders = await listFolders(client);
  const found = folders.find((f) => f.display_name.toLowerCase() === lc);
  if (!found) throw new Error(`Folder not found: ${name}`);
  return found.id;
}
