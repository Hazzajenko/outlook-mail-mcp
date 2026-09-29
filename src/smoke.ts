#!/usr/bin/env node
/**
 * Live smoke test against the real Graph API.
 *
 * The unit/integration suites feed hand-written fixtures, so they can only ever
 * confirm that the code agrees with itself. They cannot catch the case where
 * Graph returns a shape our schemas don't accept — e.g. a `$select` list that
 * stops requesting a field a Zod schema still marks required. That class of bug
 * takes down a tool completely while the suite stays green.
 *
 * This script closes that gap: it calls every exposed operation against the real
 * mailbox and parses each response through the *declared result schema*, so any
 * drift between what Graph sends and what we promise fails loudly.
 *
 * Read-only. Requires OUTLOOK_MAIL_CLIENT_ID and a cached token (`outlook-mail
 * auth`). Never runs as part of `pnpm check` — invoke it deliberately.
 */
try {
  process.loadEnvFile();
} catch {
  // no .env present; fine
}

import { z } from "zod";
import { createTokenProvider } from "./core/auth.ts";
import { resolveAuthConfig } from "./core/config.ts";
import type { GraphClient } from "./core/graph-client.ts";
import { HttpGraphClient } from "./core/http-graph-client.ts";
import {
  BriefListResultSchema,
  CountParamsSchema,
  CountResultSchema,
  FolderSchema,
  FullMessageSchema,
  ListBriefParamsSchema,
  SearchParamsSchema,
  SearchResultSchema,
} from "./core/schemas.ts";
import {
  countEmails,
  getConversation,
  getEmail,
  listEmailsBrief,
  listFolders,
  search,
} from "./core/search.ts";

/** Thrown by a case that cannot run because the mailbox lacks the data it needs. */
class Skip extends Error {}

type Outcome = "pass" | "fail" | "skip";

interface CaseResult {
  name: string;
  outcome: Outcome;
  detail: string;
}

const results: CaseResult[] = [];

async function check(name: string, fn: () => Promise<string>): Promise<void> {
  const label = name.padEnd(34);
  try {
    const detail = await fn();
    results.push({ name, outcome: "pass", detail });
    console.log(`  PASS  ${label}${detail}`);
  } catch (e) {
    if (e instanceof Skip) {
      results.push({ name, outcome: "skip", detail: e.message });
      console.log(`  SKIP  ${label}${e.message}`);
      return;
    }
    const detail = e instanceof Error ? e.message : String(e);
    results.push({ name, outcome: "fail", detail });
    console.log(`  FAIL  ${label}${summarize(detail)}`);
  }
}

/**
 * One-line teaser for the inline row; the untouched message is reprinted in the
 * failure report below. Zod messages are pretty-printed JSON, so taking the
 * first line alone would just yield "[".
 */
function summarize(s: string): string {
  const flat = s.replace(/\s+/g, " ").trim();
  return flat.length > 90 ? `${flat.slice(0, 89)}…` : flat;
}

function buildClient(): GraphClient {
  const provider = createTokenProvider(resolveAuthConfig());
  return new HttpGraphClient({ getToken: () => provider.getToken() });
}

async function main(): Promise<void> {
  const client = buildClient();
  console.log("outlook-mail smoke — live Graph, read-only\n");

  // A message id + conversation id harvested from the first search that returns
  // anything, so the single-message cases below have something real to fetch.
  let sampleId: string | undefined;
  let sampleConversationId: string | undefined;
  let sampleCursor: string | undefined;

  await check("list_folders", async () => {
    const folders = z.array(FolderSchema).parse(await listFolders(client));
    if (folders.length === 0) throw new Error("no folders returned");
    return `${folders.length} folders`;
  });

  // No free-text field -> OData $filter + $orderby path.
  await check("search (OData path)", async () => {
    const params = SearchParamsSchema.parse({ since: "-30d", top: 5 });
    const res = SearchResultSchema.parse(await search(client, params));
    if (res.results.length > 5) throw new Error(`top=5 ignored: got ${res.results.length}`);
    const first = res.results[0];
    if (first !== undefined) {
      sampleId = first.id;
      sampleConversationId = first.conversation_id;
    }
    sampleCursor = res.next_cursor;
    return `${res.total_returned} results, has_more=${res.has_more}`;
  });

  // Free-text present -> KQL $search path (no $filter, no $orderby).
  await check("search (KQL path)", async () => {
    const params = SearchParamsSchema.parse({ query: "the", top: 5 });
    const res = SearchResultSchema.parse(await search(client, params));
    return `${res.total_returned} results, has_more=${res.has_more}`;
  });

  // Folder name resolution + an OData-only filter that must not reach KQL.
  // The `since` bound is load-bearing: Graph rejects `inferenceClassification eq
  // ...` combined with `$orderby receivedDateTime desc` unless the filter also
  // restricts receivedDateTime (400 InefficientFilter). Verified against the
  // live API — `isRead eq false` + the same orderby is fine, so the limitation
  // is specific to inferenceClassification.
  await check("search (folder + inference)", async () => {
    const params = SearchParamsSchema.parse({
      folder: "inbox",
      inference_classification: "focused",
      since: "-30d",
      top: 3,
    });
    const res = SearchResultSchema.parse(await search(client, params));
    const named = res.results.filter((m) => m.folder !== undefined).length;
    return `${res.total_returned} results, ${named} with folder name`;
  });

  // The unbounded form is the failure the bound above exists to avoid. Assert we
  // turn Graph's opaque 400 into actionable advice — and that the limitation is
  // still real. If Graph starts accepting this, the case fails, which is the
  // signal to drop the translation in search.ts and the docs around it.
  await check("search (unbounded inference)", async () => {
    const params = SearchParamsSchema.parse({ inference_classification: "focused", top: 3 });
    try {
      await search(client, params);
    } catch (e) {
      const message = e instanceof Error ? e.message : String(e);
      if (!message.includes("Pass since and/or until")) {
        throw new Error(`400 reached the caller untranslated: ${message}`);
      }
      return "opaque 400 translated";
    }
    throw new Error("Graph now accepts an unbounded inference filter — drop the translation");
  });

  // The cursor path swaps the built query for a raw Graph nextLink URL — unit
  // tests fake that URL, so this is the only place it meets the real thing.
  await check("search (cursor pagination)", async () => {
    if (sampleCursor === undefined) throw new Skip("no next_cursor on page 1");
    const params = SearchParamsSchema.parse({ cursor: sampleCursor, top: 5 });
    const res = SearchResultSchema.parse(await search(client, params));
    return `${res.total_returned} results from page 2`;
  });

  await check("count_emails", async () => {
    const params = CountParamsSchema.parse({ folder: "inbox", since: "-30d" });
    const res = CountResultSchema.parse(await countEmails(client, params));
    return `count=${res.count}`;
  });

  await check("list_emails_brief", async () => {
    const params = ListBriefParamsSchema.parse({ folder: "inbox", since: "-30d", top: 5 });
    const res = BriefListResultSchema.parse(await listEmailsBrief(client, params));
    const lineCount = res.lines === "" ? 0 : res.lines.split("\n").length;
    if (lineCount !== res.total_returned) {
      throw new Error(`total_returned=${res.total_returned} but rendered ${lineCount} lines`);
    }
    return `${res.total_returned} lines, has_more=${res.has_more}`;
  });

  await check("get_email", async () => {
    if (sampleId === undefined) throw new Skip("no message id from search");
    const msg = FullMessageSchema.parse(await getEmail(client, sampleId));
    if (msg.web_link === "") throw new Error("web_link empty — FULL_SELECT no longer returns it");
    return `${msg.body.length} chars, ${msg.internet_message_headers.length} headers`;
  });

  await check("get_email (html body)", async () => {
    if (sampleId === undefined) throw new Skip("no message id from search");
    const msg = FullMessageSchema.parse(await getEmail(client, sampleId, { body_format: "html" }));
    return `content_type=${msg.body_content_type}`;
  });

  await check("get_conversation", async () => {
    if (sampleConversationId === undefined) throw new Skip("no conversation id from search");
    const thread = z
      .array(FullMessageSchema)
      .parse(await getConversation(client, sampleConversationId, { top: 10 }));
    if (thread.length === 0) throw new Error("conversation returned no messages");
    return `${thread.length} messages`;
  });

  report();
}

function report(): void {
  const passed = results.filter((r) => r.outcome === "pass").length;
  const failed = results.filter((r) => r.outcome === "fail");
  const skipped = results.filter((r) => r.outcome === "skip").length;

  console.log(`\n${passed} passed, ${failed.length} failed, ${skipped} skipped`);

  if (failed.length > 0) {
    console.log("\nFailures:");
    for (const f of failed) {
      console.log(`\n  ${f.name}`);
      for (const line of f.detail.split("\n")) console.log(`    ${line}`);
    }
    process.exitCode = 1;
  }
}

main().catch((e: unknown) => {
  console.error(e instanceof Error ? e.message : String(e));
  process.exitCode = 1;
});
