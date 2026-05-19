import { z } from "zod";
import type { GraphClient } from "./graph-client.ts";
import {
  buildCountQuery,
  buildGraphQuery,
  CONVERSATION_SELECT,
  FULL_SELECT,
} from "./query-builder.ts";
import { mapFolder, mapFullMessage, mapLeanWithFolderId } from "./result-mapper.ts";
import type {
  CountParams,
  CountResult,
  Folder,
  FullMessage,
  LeanMessage,
  SearchParams,
  SearchResult,
} from "./schemas.ts";

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

  const out: LeanMessage[] = [];
  const folderIds: (string | undefined)[] = [];
  let nextCursor: string | undefined;
  let hasMore = false;

  while (true) {
    const raw = await client.get(pagePath, pageQuery);
    const { value, "@odata.nextLink": nextLink } = PageSchema.parse(raw);
    const remaining = params.top - out.length;

    if (value.length > remaining) {
      // truncating mid-page: nextLink would skip unconsumed items, so don't expose a cursor
      for (let i = 0; i < remaining; i++) {
        const { lean, folder_id } = mapLeanWithFolderId(value[i]);
        out.push(lean);
        folderIds.push(folder_id);
      }
      hasMore = true;
      break;
    }

    for (const v of value) {
      const { lean, folder_id } = mapLeanWithFolderId(v);
      out.push(lean);
      folderIds.push(folder_id);
    }

    if (out.length >= params.top) {
      nextCursor = nextLink;
      hasMore = nextLink !== undefined;
      break;
    }
    if (nextLink === undefined) break;

    pagePath = nextLink;
    pageQuery = undefined;
  }

  await resolveFolderNames(client, out, folderIds);

  return {
    results: out,
    total_returned: out.length,
    has_more: hasMore,
    ...(nextCursor !== undefined ? { next_cursor: nextCursor } : {}),
  };
}

async function resolveFolderNames(
  client: GraphClient,
  messages: LeanMessage[],
  folderIds: (string | undefined)[],
): Promise<void> {
  if (folderIds.every((id) => id === undefined)) return;
  const folders = await listFolders(client);
  const byId = new Map(folders.map((f) => [f.id, f.display_name]));
  for (let i = 0; i < messages.length; i++) {
    const id = folderIds[i];
    const m = messages[i];
    if (id === undefined || m === undefined) continue;
    const name = byId.get(id);
    if (name !== undefined) m.folder = name;
  }
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

export interface GetConversationOptions {
  body_format?: "text" | "html";
  top?: number;
}

export async function getConversation(
  client: GraphClient,
  conversationId: string,
  opts: GetConversationOptions = {},
): Promise<FullMessage[]> {
  const top = opts.top ?? 200;
  const format = opts.body_format ?? "text";
  const escaped = conversationId.replace(/'/g, "''");

  const query = new URLSearchParams();
  query.set("$top", String(top));
  query.set("$select", CONVERSATION_SELECT);
  query.set("$filter", `conversationId eq '${escaped}'`);

  const headers = { Prefer: `outlook.body-content-type="${format}"` };

  const out: FullMessage[] = [];
  let pagePath: string = "/me/messages";
  let pageQuery: URLSearchParams | undefined = query;

  while (out.length < top) {
    const raw = await client.get(pagePath, pageQuery, headers);
    const { value, "@odata.nextLink": nextLink } = PageSchema.parse(raw);
    const remaining = top - out.length;
    const take = Math.min(value.length, remaining);
    for (let i = 0; i < take; i++) out.push(mapFullMessage(value[i]));
    if (out.length >= top) break;
    if (nextLink === undefined) break;
    pagePath = nextLink;
    pageQuery = undefined;
  }

  out.sort((a, b) => a.received_at.localeCompare(b.received_at));
  return out;
}

const CountResponseSchema = z.object({
  "@odata.count": z.number().int(),
});

export async function countEmails(client: GraphClient, params: CountParams): Promise<CountResult> {
  const folderId = params.folder ? await resolveFolderId(client, params.folder) : undefined;
  const { endpoint, query } = buildCountQuery(params, folderId);
  // Graph requires this header for $count with $search; harmless for $filter-only.
  const raw = await client.get(endpoint, query, { ConsistencyLevel: "eventual" });
  const { "@odata.count": count } = CountResponseSchema.parse(raw);
  return { count };
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
