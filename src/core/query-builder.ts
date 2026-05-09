import type { SearchParams } from "./schemas.ts";

export const LEAN_SELECT = [
  "id",
  "from",
  "subject",
  "receivedDateTime",
  "hasAttachments",
  "isRead",
  "conversationId",
  "webLink",
  "bodyPreview",
  "inferenceClassification",
  "parentFolderId",
].join(",");

export const FULL_SELECT = [
  ...LEAN_SELECT.split(","),
  "toRecipients",
  "ccRecipients",
  "body",
  "importance",
  "internetMessageHeaders",
].join(",");

export const CONVERSATION_SELECT = [
  ...LEAN_SELECT.split(","),
  "toRecipients",
  "ccRecipients",
  "body",
  "importance",
].join(",");

export interface GraphQuery {
  endpoint: string;
  query: URLSearchParams;
}

export function buildGraphQuery(params: SearchParams, folderId?: string): GraphQuery {
  const endpoint = folderId ? `/me/mailFolders/${folderId}/messages` : "/me/messages";

  const query = new URLSearchParams();
  query.set("$top", String(params.top));
  query.set("$select", LEAN_SELECT);

  if (hasFreeText(params)) {
    if (params.inference_classification) {
      throw new Error(
        "inference_classification cannot be combined with text search (query/from/to/subject_contains/body_contains); Graph KQL has no inference keyword",
      );
    }
    query.set("$search", `"${buildKql(params)}"`);
  } else {
    query.set("$orderby", "receivedDateTime desc");
    const filter = buildOdataFilter(params);
    if (filter) {
      query.set("$filter", filter);
    }
  }

  return { endpoint, query };
}

function hasFreeText(p: SearchParams): boolean {
  return (
    p.query !== undefined ||
    p.from !== undefined ||
    p.to !== undefined ||
    p.subject_contains !== undefined ||
    p.body_contains !== undefined
  );
}

function buildKql(p: SearchParams): string {
  const terms: string[] = [];
  if (p.query) terms.push(p.query);
  if (p.from) terms.push(`from:${p.from}`);
  if (p.to) terms.push(`to:${p.to}`);
  if (p.subject_contains) terms.push(`subject:${p.subject_contains}`);
  if (p.body_contains) terms.push(`body:${p.body_contains}`);
  if (p.has_attachment !== undefined) {
    terms.push(`hasattachment:${p.has_attachment ? "yes" : "no"}`);
  }
  if (p.is_unread !== undefined) {
    terms.push(`read:${p.is_unread ? "no" : "yes"}`);
  }
  if (p.importance) terms.push(`importance:${p.importance}`);
  if (p.since) terms.push(`received>=${dateOnly(p.since)}`);
  if (p.until) terms.push(`received<=${dateOnly(p.until)}`);
  return terms.join(" ");
}

function buildOdataFilter(p: SearchParams): string {
  const parts: string[] = [];
  if (p.since) parts.push(`receivedDateTime ge ${p.since.toISOString()}`);
  if (p.until) parts.push(`receivedDateTime le ${p.until.toISOString()}`);
  if (p.is_unread !== undefined) parts.push(`isRead eq ${!p.is_unread}`);
  if (p.has_attachment !== undefined) parts.push(`hasAttachments eq ${p.has_attachment}`);
  if (p.importance) parts.push(`importance eq '${p.importance}'`);
  if (p.inference_classification) {
    parts.push(`inferenceClassification eq '${p.inference_classification}'`);
  }
  return parts.join(" and ");
}

function dateOnly(d: Date): string {
  return d.toISOString().slice(0, 10);
}
