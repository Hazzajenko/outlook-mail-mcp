#!/usr/bin/env node
try {
  process.loadEnvFile();
} catch {
  // no .env present; fine
}

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { createTokenProvider } from "./core/auth.ts";
import { resolveAuthConfig } from "./core/config.ts";
import type { GraphClient } from "./core/graph-client.ts";
import { HttpGraphClient } from "./core/http-graph-client.ts";
import { CountParamsSchema, ListBriefParamsSchema, SearchParamsSchema } from "./core/schemas.ts";
import {
  countEmails,
  getConversation,
  getEmail,
  listEmailsBrief,
  listFolders,
  search,
} from "./core/search.ts";
import { VERSION } from "./version.ts";

let cachedClient: GraphClient | undefined;
function client(): GraphClient {
  if (cachedClient !== undefined) return cachedClient;
  // interactive: false — a device-code prompt inside an MCP tool call is
  // invisible to the user (stderr goes to the client's logs) and hangs the
  // call; fail fast and direct them to `outlook-mail auth` instead.
  const tokenProvider = createTokenProvider({ ...resolveAuthConfig(), interactive: false });
  cachedClient = new HttpGraphClient({ getToken: () => tokenProvider.getToken() });
  return cachedClient;
}

const server = new McpServer({
  name: "outlook-mail",
  version: VERSION,
});

server.registerTool(
  "search_emails",
  {
    title: "Search Outlook emails",
    description: [
      "Search the user's Outlook mailbox via Microsoft Graph. Returns lean message metadata + body preview; use get_email for full body.",
      "",
      'FOLDER SCOPE: Without `folder`, Graph $search excludes Junk Email, Deleted Items, and Drafts by default. To search those, pass `folder: "junkemail"` (or `deleteditems`/`drafts`). Use list_folders to discover custom folder names.',
      "",
      "TWO QUERY PATHS (auto-selected):",
      "  • Text path — triggered by query/from/to/subject_contains/body_contains. Uses KQL $search; date precision is per-day; results NOT sorted by date.",
      "  • Structured path — no text params. Uses OData $filter; sorted by receivedDateTime desc; supports inference_classification.",
      "",
      "MUTUAL EXCLUSION: `inference_classification` cannot combine with any text param (KQL has no inference keyword).",
      "",
      "DATE BOUNDS: `since` is inclusive. `until` with a bare date (2026-05-08) includes the entire day; with an explicit time it's inclusive of that instant.",
      "",
      "PAGINATION: Default top=50, max 500. If `has_more: true`, fetch more by passing the response's `next_cursor` back as `cursor` (other params are ignored when cursor is set). Mid-page truncation can set has_more=true without a cursor — re-issue with a higher `top` to recover.",
    ].join("\n"),
    inputSchema: SearchParamsSchema.shape,
  },
  async (args) => {
    const result = await search(client(), args);
    return {
      content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
    };
  },
);

server.registerTool(
  "get_email",
  {
    title: "Get full email by id",
    description:
      "Fetch the full body of one email by its Graph id. Body returned as plain text by default (server-side HTML→text conversion); pass body_format='html' for raw HTML. Returns only notable internet headers (Authentication-Results, Return-Path, Reply-To, List-Unsubscribe); the ~60 Graph sends are mostly spam-filter internals that cost more tokens than the body. Pass include_all_headers=true when debugging mail routing or authenticity.",
    inputSchema: {
      id: z.string().describe("Graph message id"),
      body_format: z.enum(["text", "html"]).optional().describe("Body format; defaults to 'text'"),
      include_all_headers: z
        .boolean()
        .optional()
        .describe(
          "Return every internet message header instead of the notable few. Verbose — only for routing/authenticity debugging.",
        ),
    },
  },
  async ({ id, body_format, include_all_headers }) => {
    const result = await getEmail(client(), id, {
      ...(body_format ? { body_format } : {}),
      ...(include_all_headers !== undefined ? { include_all_headers } : {}),
    });
    return {
      content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
    };
  },
);

server.registerTool(
  "get_conversation",
  {
    title: "Get full email thread by conversation_id",
    description:
      "Fetch all messages in a conversation/thread, sorted oldest-first, with bodies. Use the conversation_id from search_emails results. Body returned as plain text by default; pass body_format='html' for raw HTML. Capped at 200 messages by default.",
    inputSchema: {
      conversation_id: z.string().describe("Graph conversationId from a search_emails result"),
      body_format: z.enum(["text", "html"]).optional().describe("Body format; defaults to 'text'"),
      top: z
        .number()
        .int()
        .min(1)
        .max(1000)
        .optional()
        .describe("Max messages to return; defaults to 200"),
    },
  },
  async ({ conversation_id, body_format, top }) => {
    const opts: { body_format?: "text" | "html"; top?: number } = {};
    if (body_format) opts.body_format = body_format;
    if (top !== undefined) opts.top = top;
    const result = await getConversation(client(), conversation_id, opts);
    return {
      content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
    };
  },
);

server.registerTool(
  "count_emails",
  {
    title: "Count Outlook emails matching a filter",
    description: [
      "Return the total count of emails matching the filter. No bodies fetched; cheap aggregate.",
      "",
      "Same filter params as search_emails (folder, since/until, from, query, is_unread, inference_classification, etc.). top/cursor not applicable.",
      "",
      "USE BEFORE BULK OPERATIONS: e.g. count first to decide whether to call list_emails_brief (cheap if <2000) or to narrow filters further.",
    ].join("\n"),
    inputSchema: CountParamsSchema.shape,
  },
  async (args) => {
    const result = await countEmails(client(), args);
    return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
  },
);

server.registerTool(
  "list_emails_brief",
  {
    title: "List Outlook emails in compact text format",
    description: [
      "Return a compact text list — one line per email — for bulk scanning a date range. Format: `YYYY-MM-DD HH:MM | sender@domain | subject` (subject truncated to 80 chars).",
      "",
      "Same filter params as search_emails. Default top=500, max 2000. Auto-paginates Graph; no cursor exposed.",
      "",
      'USE FOR: "look through my last month", inbox-shape questions, sender breakdown. NOT for full bodies (use get_email) or pagination workflows (use search_emails).',
      "",
      "TOKEN BUDGET: ~70-130 chars per email. 500 emails ≈ 10K tokens, 1000 ≈ 20K, 2000 ≈ 40K. Run count_emails first to gauge.",
    ].join("\n"),
    inputSchema: ListBriefParamsSchema.shape,
  },
  async (args) => {
    const result = await listEmailsBrief(client(), args);
    return { content: [{ type: "text", text: JSON.stringify(result, null, 2) }] };
  },
);

server.registerTool(
  "list_folders",
  {
    title: "List mail folders",
    description:
      "List the user's mail folders (Inbox, Sent, custom folders, etc.) with item counts. Use to discover folder names for search_emails.",
    inputSchema: {},
  },
  async () => {
    const result = await listFolders(client());
    return {
      content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
    };
  },
);

const transport = new StdioServerTransport();
await server.connect(transport);
