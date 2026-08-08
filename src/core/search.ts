import { z } from "zod";
import { type GraphClient, GraphHttpError } from "./graph-client.ts";
import {
  buildBriefQuery,
  buildCountQuery,
  buildGraphQuery,
  CONVERSATION_SELECT,
  FULL_SELECT,
} from "./query-builder.ts";
import { mapFolder, mapFullMessage, mapLeanWithFolderId } from "./result-mapper.ts";
import type {
  BriefListResult,
  CountParams,
  CountResult,
  FilterParams,
  Folder,
  FullMessage,
  LeanMessage,
  ListBriefParams,
  SearchParams,
  SearchResult,
} from "./schemas.ts";

const PageSchema = z.object({
  value: z.array(z.unknown()),
  "@odata.nextLink": z.string().optional(),
});

/**
 * Graph rejects an `inferenceClassification` filter that isn't bounded by a
 * receivedDateTime range when the results are also sorted by receivedDateTime,
 * with a bare 400 InefficientFilter that names neither the cause nor the fix.
 * Translate it here, where the params that caused it are still in scope.
 *
 * Only the sorted operations are affected — countEmails builds its query with
 * withOrderBy: false and works unbounded, so it does not use this wrapper.
 */
async function getSorted(
  client: GraphClient,
  params: FilterParams,
  pathOrUrl: string,
  query?: URLSearchParams,
): Promise<unknown> {
  try {
    return await client.get(pathOrUrl, query);
  } catch (e) {
    if (
      e instanceof GraphHttpError &&
      e.body.includes("InefficientFilter") &&
      params.inference_classification !== undefined &&
      params.since === undefined &&
      params.until === undefined
    ) {
      throw new Error(
        `Graph rejected inference_classification='${params.inference_classification}' with 400 InefficientFilter: ` +
          "results are sorted by receivedDateTime and Graph cannot combine that sort with an unbounded " +
          "inference filter. Pass since and/or until to bound the range. (count_emails does not sort, " +
          "so it works without a bound.)",
        { cause: e },
      );
    }
    throw e;
  }
}

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
    const raw = await getSorted(client, params, pagePath, pageQuery);
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

/**
 * Headers worth returning by default.
 *
 * Graph sends ~60 per message — Exchange spam-filter internals, DKIM
 * signatures, `X-Microsoft-Antispam-Message-Info` blobs — which routinely cost
 * several times more tokens than the body they arrived with. These four are the
 * ones a reader acts on: the first two are the anti-spoofing verdict
 * `renderFullMessage` already prints, `Reply-To` frequently differs from `From`
 * and is often the only real human address on an automated mail, and
 * `List-Unsubscribe` answers "how do I stop this".
 *
 * Graph has no server-side way to select a subset, so this trims on receipt —
 * it saves tokens, not bytes on the wire. Pass `include_all_headers` for the
 * full set when actually debugging mail routing.
 */
const NOTABLE_HEADERS = new Set(
  ["Authentication-Results", "Return-Path", "Reply-To", "List-Unsubscribe"].map((h) =>
    h.toLowerCase(),
  ),
);

export interface GetEmailOptions {
  body_format?: "text" | "html";
  include_all_headers?: boolean;
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
  const message = mapFullMessage(raw);
  if (opts.include_all_headers) return message;
  // Senders capitalise header names inconsistently; match on the wire name.
  message.internet_message_headers = message.internet_message_headers.filter((h) =>
    NOTABLE_HEADERS.has(h.name.toLowerCase()),
  );
  return message;
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

const BriefMessageSchema = z.object({
  subject: z.string().nullable().optional(),
  receivedDateTime: z.string(),
  from: z
    .object({ emailAddress: z.object({ address: z.string() }) })
    .nullable()
    .optional(),
});

export async function listEmailsBrief(
  client: GraphClient,
  params: ListBriefParams,
): Promise<BriefListResult> {
  const folderId = params.folder ? await resolveFolderId(client, params.folder) : undefined;
  const built = buildBriefQuery(params, folderId);
  let pagePath: string = built.endpoint;
  let pageQuery: URLSearchParams | undefined = built.query;

  const lines: string[] = [];
  let hasMore = false;

  while (true) {
    const raw = await getSorted(client, params, pagePath, pageQuery);
    const { value, "@odata.nextLink": nextLink } = PageSchema.parse(raw);
    const remaining = params.top - lines.length;

    if (value.length > remaining) {
      for (let i = 0; i < remaining; i++) lines.push(formatBriefLine(value[i]));
      hasMore = true;
      break;
    }

    for (const v of value) lines.push(formatBriefLine(v));

    if (lines.length >= params.top) {
      hasMore = nextLink !== undefined;
      break;
    }
    if (nextLink === undefined) break;

    pagePath = nextLink;
    pageQuery = undefined;
  }

  return { lines: lines.join("\n"), total_returned: lines.length, has_more: hasMore };
}

function formatBriefLine(raw: unknown): string {
  const m = BriefMessageSchema.parse(raw);
  const date = formatBriefDate(m.receivedDateTime);
  const from = m.from?.emailAddress.address ?? "(unknown)";
  const subject = truncate((m.subject ?? "").replace(/\s+/g, " ").trim(), 80);
  return `${date} | ${from} | ${subject}`;
}

function formatBriefDate(iso: string): string {
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return iso;
  const y = d.getUTCFullYear();
  const mo = String(d.getUTCMonth() + 1).padStart(2, "0");
  const dd = String(d.getUTCDate()).padStart(2, "0");
  const hh = String(d.getUTCHours()).padStart(2, "0");
  const mi = String(d.getUTCMinutes()).padStart(2, "0");
  return `${y}-${mo}-${dd} ${hh}:${mi}`;
}

function truncate(s: string, n: number): string {
  return s.length > n ? `${s.slice(0, n - 1)}…` : s;
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
