import type {
  EmailAddress,
  Folder,
  FullMessage,
  LeanMessage,
  SearchResult,
} from "./core/schemas.ts";

const DIVIDER = "─".repeat(60);

export function renderSearchResults(result: SearchResult): string {
  if (result.results.length === 0) {
    return "0 results.\n";
  }
  const lines: string[] = [];
  result.results.forEach((m, i) => {
    lines.push(`[${i + 1}] ${formatLeanHeader(m)}`);
    lines.push(`    ${m.subject || "(no subject)"}`);
    if (m.body_preview) lines.push(`    ${truncate(m.body_preview, 120)}`);
    lines.push("");
  });
  const footer = result.more_available
    ? `${result.total_returned} results · more available`
    : `${result.total_returned} results`;
  lines.push(footer);
  return `${lines.join("\n")}\n`;
}

const AUTH_HEADERS = ["Authentication-Results", "Return-Path"];

export function renderFullMessage(m: FullMessage): string {
  const lines: string[] = [];
  lines.push(`From:    ${formatAddress(m.from)}`);
  if (m.to.length > 0) lines.push(`To:      ${m.to.map(formatAddress).join(", ")}`);
  if (m.cc.length > 0) lines.push(`Cc:      ${m.cc.map(formatAddress).join(", ")}`);
  lines.push(`Subject: ${m.subject || "(no subject)"}`);
  lines.push(`Date:    ${formatDate(m.received_at)}`);
  if (m.importance !== "normal") lines.push(`Importance: ${m.importance}`);
  if (m.inference_classification === "other") lines.push("Classification: other");
  if (m.has_attachment) lines.push("Attachments: yes");
  if (m.body_content_type === "html") {
    lines.push("Body type: html (raw)");
  }
  for (const name of AUTH_HEADERS) {
    const h = m.internet_message_headers.find((x) => x.name === name);
    if (h) lines.push(`${h.name}: ${h.value}`);
  }
  lines.push(DIVIDER);
  lines.push(m.body);
  return `${lines.join("\n")}\n`;
}

export function renderFolders(folders: Folder[]): string {
  if (folders.length === 0) return "0 folders.\n";
  const nameWidth = Math.max(...folders.map((f) => f.display_name.length));
  const lines = folders.map((f) => {
    const padded = f.display_name.padEnd(nameWidth);
    return `${padded}  ${f.total_item_count} items, ${f.unread_item_count} unread`;
  });
  return `${lines.join("\n")}\n`;
}

function formatLeanHeader(m: LeanMessage): string {
  const date = formatDate(m.received_at);
  const flags: string[] = [];
  if (!m.is_read) flags.push("●unread");
  if (m.has_attachment) flags.push("📎");
  const flagStr = flags.length > 0 ? ` ${flags.join(" ")}` : "";
  return `${date}  ${formatAddress(m.from)}${flagStr}`;
}

function formatAddress(a: EmailAddress): string {
  return a.name ? `${a.name} <${a.address}>` : a.address;
}

function formatDate(iso: string): string {
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return iso;
  const yyyy = d.getUTCFullYear();
  const mm = String(d.getUTCMonth() + 1).padStart(2, "0");
  const dd = String(d.getUTCDate()).padStart(2, "0");
  const hh = String(d.getUTCHours()).padStart(2, "0");
  const mi = String(d.getUTCMinutes()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd} ${hh}:${mi}`;
}

function truncate(s: string, max: number): string {
  if (s.length <= max) return s;
  return `${s.slice(0, max - 1)}…`;
}
