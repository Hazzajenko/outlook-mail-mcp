#!/usr/bin/env node
try {
  process.loadEnvFile();
} catch {
  // no .env present; fine
}

import { Command } from "commander";
import { z } from "zod";
import { renderFolders, renderFullMessage, renderSearchResults } from "./cli-render.ts";
import { createTokenProvider, type TokenProvider } from "./core/auth.ts";
import type { GraphClient } from "./core/graph-client.ts";
import { HttpGraphClient } from "./core/http-graph-client.ts";
import { type SearchParamsInput, SearchParamsSchema } from "./core/schemas.ts";
import { getConversation, getEmail, listFolders, search } from "./core/search.ts";

const SearchOptsSchema = z.object({
  query: z.string().optional(),
  from: z.string().optional(),
  to: z.string().optional(),
  subject: z.string().optional(),
  body: z.string().optional(),
  since: z.string().optional(),
  until: z.string().optional(),
  hasAttachment: z.boolean().optional(),
  unread: z.boolean().optional(),
  read: z.boolean().optional(),
  folder: z.string().optional(),
  importance: z.string().optional(),
  inferenceClassification: z.string().optional(),
  top: z.number().int().optional(),
  json: z.boolean().optional(),
});

type SearchOpts = z.infer<typeof SearchOptsSchema>;

function buildTokenProvider(): TokenProvider {
  const clientId = process.env.OUTLOOK_QUERY_CLIENT_ID;
  if (!clientId) {
    throw new Error(
      "OUTLOOK_QUERY_CLIENT_ID env var not set. Register an Azure app and export the client ID.",
    );
  }
  const tenantId = process.env.OUTLOOK_QUERY_TENANT_ID;
  return createTokenProvider(tenantId !== undefined ? { clientId, tenantId } : { clientId });
}

function buildClient(): GraphClient {
  const provider = buildTokenProvider();
  return new HttpGraphClient({ getToken: () => provider.getToken() });
}

function optsToSearchParams(opts: SearchOpts): SearchParamsInput {
  const raw: SearchParamsInput = {};
  if (opts.query !== undefined) raw.query = opts.query;
  if (opts.from !== undefined) raw.from = opts.from;
  if (opts.to !== undefined) raw.to = opts.to;
  if (opts.subject !== undefined) raw.subject_contains = opts.subject;
  if (opts.body !== undefined) raw.body_contains = opts.body;
  if (opts.since !== undefined) raw.since = opts.since;
  if (opts.until !== undefined) raw.until = opts.until;
  if (opts.hasAttachment !== undefined) raw.has_attachment = opts.hasAttachment;
  if (opts.unread === true) raw.is_unread = true;
  else if (opts.read === true) raw.is_unread = false;
  if (opts.folder !== undefined) raw.folder = opts.folder;
  if (opts.importance !== undefined) {
    raw.importance = opts.importance as "low" | "normal" | "high";
  }
  if (opts.inferenceClassification !== undefined) {
    raw.inference_classification = opts.inferenceClassification as "focused" | "other";
  }
  if (opts.top !== undefined) raw.top = opts.top;
  return raw;
}

const program = new Command();

program
  .name("outlook-query")
  .description("Query personal Outlook via Microsoft Graph")
  .version("0.0.0");

program
  .command("search")
  .description("Search emails")
  .option("-q, --query <text>", "free-text search")
  .option("--from <addr>", "sender address or domain")
  .option("--to <addr>", "recipient address or domain")
  .option("--subject <text>", "subject contains")
  .option("--body <text>", "body contains")
  .option("--since <when>", "ISO date or relative (-7d, -2w, -3h)")
  .option("--until <when>", "ISO date or relative")
  .option("--has-attachment", "only with attachments")
  .option("--unread", "only unread")
  .option("--read", "only read (already-opened)")
  .option("--folder <name>", "folder (well-known: inbox/sent/archive/... or custom name)")
  .option("--importance <level>", "low | normal | high")
  .option(
    "--inference-classification <value>",
    "focused | other (Outlook Focused/Other split); needs --since/--until",
  )
  .option("--top <n>", "max results (1-500, default 50)", (v) => parseInt(v, 10))
  .option("--json", "JSON output")
  .action(async (rawOpts: unknown) => {
    const opts = SearchOptsSchema.parse(rawOpts);
    const params = SearchParamsSchema.parse(optsToSearchParams(opts));
    const client = buildClient();
    const result = await search(client, params);
    if (opts.json) {
      process.stdout.write(`${JSON.stringify(result, null, 2)}\n`);
    } else {
      process.stdout.write(renderSearchResults(result));
    }
  });

program
  .command("get <id>")
  .description("Get full email by id")
  .option("--body-format <fmt>", "text | html (default: text)")
  .option("--json", "JSON output")
  .action(async (id: string, rawOpts: unknown) => {
    const opts = z
      .object({
        json: z.boolean().optional(),
        bodyFormat: z.enum(["text", "html"]).optional(),
      })
      .parse(rawOpts);
    const client = buildClient();
    const result = await getEmail(
      client,
      id,
      opts.bodyFormat ? { body_format: opts.bodyFormat } : {},
    );
    if (opts.json) {
      process.stdout.write(`${JSON.stringify(result, null, 2)}\n`);
    } else {
      process.stdout.write(renderFullMessage(result));
    }
  });

program
  .command("conversation <id>")
  .description("Get full email thread by conversation_id")
  .option("--body-format <fmt>", "text | html (default: text)")
  .option("--top <n>", "max messages (default 200)", (v) => parseInt(v, 10))
  .option("--json", "JSON output")
  .action(async (id: string, rawOpts: unknown) => {
    const opts = z
      .object({
        json: z.boolean().optional(),
        bodyFormat: z.enum(["text", "html"]).optional(),
        top: z.number().int().optional(),
      })
      .parse(rawOpts);
    const client = buildClient();
    const convoOpts: { body_format?: "text" | "html"; top?: number } = {};
    if (opts.bodyFormat) convoOpts.body_format = opts.bodyFormat;
    if (opts.top !== undefined) convoOpts.top = opts.top;
    const messages = await getConversation(client, id, convoOpts);
    if (opts.json) {
      process.stdout.write(`${JSON.stringify(messages, null, 2)}\n`);
    } else {
      for (const m of messages) process.stdout.write(renderFullMessage(m));
    }
  });

program
  .command("folders")
  .description("List mail folders")
  .option("--json", "JSON output")
  .action(async (rawOpts: unknown) => {
    const opts = z.object({ json: z.boolean().optional() }).parse(rawOpts);
    const client = buildClient();
    const folders = await listFolders(client);
    if (opts.json) {
      process.stdout.write(`${JSON.stringify(folders, null, 2)}\n`);
    } else {
      process.stdout.write(renderFolders(folders));
    }
  });

program
  .command("auth")
  .description("Trigger device-code auth (populates token cache)")
  .action(async () => {
    const provider = buildTokenProvider();
    await provider.getToken();
    process.stderr.write("✓ token acquired and cached\n");
  });

program.parseAsync(process.argv).catch((e: unknown) => {
  process.stderr.write(`Error: ${(e as Error).message}\n`);
  process.exit(1);
});
