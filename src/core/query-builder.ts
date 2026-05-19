import type { FilterParams, SearchParams } from "./schemas.ts";

export const LEAN_SELECT = [
  "id",
  "from",
  "subject",
  "receivedDateTime",
  "hasAttachments",
  "isRead",
  "conversationId",
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
  "webLink",
].join(",");

export const CONVERSATION_SELECT = [
  ...LEAN_SELECT.split(","),
  "toRecipients",
  "ccRecipients",
  "body",
  "importance",
  "webLink",
].join(",");

export interface GraphQuery {
  endpoint: string;
  query: URLSearchParams;
}

function endpointFor(folderId: string | undefined): string {
  return folderId ? `/me/mailFolders/${folderId}/messages` : "/me/messages";
}

function applyFilters(query: URLSearchParams, params: FilterParams, withOrderBy: boolean): void {
  if (hasFreeText(params)) {
    if (params.inference_classification) {
      throw new Error(
        "inference_classification cannot be combined with text search (query/from/to/subject_contains/body_contains); Graph KQL has no inference keyword",
      );
    }
    query.set("$search", `"${buildKql(params)}"`);
  } else {
    if (withOrderBy) query.set("$orderby", "receivedDateTime desc");
    const filter = buildOdataFilter(params);
    if (filter) query.set("$filter", filter);
  }
}

export function buildGraphQuery(params: SearchParams, folderId?: string): GraphQuery {
  const query = new URLSearchParams();
  query.set("$top", String(params.top));
  query.set("$select", LEAN_SELECT);
  applyFilters(query, params, true);
  return { endpoint: endpointFor(folderId), query };
}

export function buildCountQuery(params: FilterParams, folderId?: string): GraphQuery {
  const query = new URLSearchParams();
  query.set("$count", "true");
  query.set("$top", "1");
  query.set("$select", "id");
  applyFilters(query, params, false);
  return { endpoint: endpointFor(folderId), query };
}

function hasFreeText(p: FilterParams): boolean {
  return (
    p.query !== undefined ||
    p.from !== undefined ||
    p.to !== undefined ||
    p.subject_contains !== undefined ||
    p.body_contains !== undefined
  );
}

function buildKql(p: FilterParams): string {
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
  if (p.since) terms.push(`received>=${dateOnly(p.since.date)}`);
  if (p.until) {
    // KQL `received` is date-precision; use exclusive `<` against the day after
    // p.until so the named day is fully included regardless of input precision.
    terms.push(`received<${dateOnly(nextDayUtc(p.until.date))}`);
  }
  return terms.join(" ");
}

function buildOdataFilter(p: FilterParams): string {
  const parts: string[] = [];
  if (p.since) parts.push(`receivedDateTime ge ${p.since.date.toISOString()}`);
  if (p.until) {
    if (p.until.dateOnly) {
      // include the whole day: < start of next day
      parts.push(`receivedDateTime lt ${nextDayUtc(p.until.date).toISOString()}`);
    } else {
      parts.push(`receivedDateTime le ${p.until.date.toISOString()}`);
    }
  }
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

function nextDayUtc(d: Date): Date {
  return new Date(d.getTime() + 86_400_000);
}
