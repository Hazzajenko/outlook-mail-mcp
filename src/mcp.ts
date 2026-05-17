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
import type { GraphClient } from "./core/graph-client.ts";
import { HttpGraphClient } from "./core/http-graph-client.ts";
import { SearchParamsSchema } from "./core/schemas.ts";
import { getConversation, getEmail, listFolders, search } from "./core/search.ts";

let cachedClient: GraphClient | undefined;
function client(): GraphClient {
  if (cachedClient !== undefined) return cachedClient;
  const clientId = process.env.OUTLOOK_QUERY_CLIENT_ID;
  if (!clientId) {
    throw new Error("OUTLOOK_QUERY_CLIENT_ID env var not set.");
  }
  const tenantId = process.env.OUTLOOK_QUERY_TENANT_ID;
  const tokenProvider = createTokenProvider(
    tenantId !== undefined ? { clientId, tenantId } : { clientId },
  );
  cachedClient = new HttpGraphClient({ getToken: () => tokenProvider.getToken() });
  return cachedClient;
}

const server = new McpServer({
  name: "outlook-query",
  version: "0.0.0",
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
      "Fetch the full body and headers of one email by its Graph id. Body returned as plain text by default (server-side HTML→text conversion); pass body_format='html' for raw HTML.",
    inputSchema: {
      id: z.string().describe("Graph message id"),
      body_format: z.enum(["text", "html"]).optional().describe("Body format; defaults to 'text'"),
    },
  },
  async ({ id, body_format }) => {
    const result = await getEmail(client(), id, body_format ? { body_format } : {});
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
