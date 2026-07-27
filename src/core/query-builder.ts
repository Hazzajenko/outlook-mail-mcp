import type { FilterParams, ListBriefParams, SearchParams } from "./schemas.ts";

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

export const BRIEF_SELECT = ["from", "subject", "receivedDateTime"].join(",");

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

export function buildBriefQuery(params: ListBriefParams, folderId?: string): GraphQuery {
  const query = new URLSearchParams();
  // Graph caps per-page at 1000; we'll auto-paginate up to params.top
  query.set("$top", String(Math.min(params.top, 1000)));
  query.set("$select", BRIEF_SELECT);
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

/**
 * Escape a value for transport inside the double-quoted `$search` string. Graph
 * accepts `\"` there; a bare `"` is a 400 ("Syntax error: character '\"' is not
 * valid at position N"). Escaping preserves KQL phrase semantics rather than
 * destroying them — Graph unwraps the transport string first, so `\"a b\"`
 * still reaches KQL as the phrase `"a b"`.
 */
function escapeKql(v: string): string {
  return v.replace(/\\/g, "\\\\").replace(/"/g, '\\"');
}

/**
 * Build a field-scoped KQL term.
 *
 * These params are documented as substring matches, not KQL, so a multi-word
 * value has to be a phrase. Bare interpolation degrades silently: KQL reads
 * `subject:online assessment` as `subject:online` AND a loose `assessment`,
 * which matches mail with neither word in the subject. Phrase-wrapping binds
 * the whole value to the field.
 *
 * Inner quotes are dropped rather than escaped — KQL has no in-phrase escape
 * for them, so `\"say \"hi\" there\"` would terminate the phrase early.
 */
function kqlField(field: string, value: string): string {
  const v = value.replace(/"/g, " ").trim();
  const esc = escapeKql(v);
  return /\s/.test(v) ? `${field}:\\"${esc}\\"` : `${field}:${esc}`;
}

function buildKql(p: FilterParams): string {
  const terms: string[] = [];
  if (p.query) {
    // `query` is documented as raw KQL, so the caller's own quotes are theirs to
    // spend on phrases. An odd count can only produce an unterminated phrase, so
    // name it here instead of letting Graph answer with an opaque 400.
    if (((p.query.match(/"/g) ?? []).length & 1) === 1) {
      throw new Error(
        `query has an unbalanced double quote: ${p.query}. Quotes in query delimit KQL phrases — pair them, or drop them to search the words separately.`,
      );
    }
    terms.push(escapeKql(p.query));
  }
  if (p.from) terms.push(kqlField("from", p.from));
  if (p.to) terms.push(kqlField("to", p.to));
  if (p.subject_contains) terms.push(kqlField("subject", p.subject_contains));
  if (p.body_contains) terms.push(kqlField("body", p.body_contains));
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
