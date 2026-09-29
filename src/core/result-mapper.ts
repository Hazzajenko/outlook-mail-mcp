import { z } from "zod";
import {
  type EmailAddress,
  type Folder,
  type FullMessage,
  ImportanceSchema,
  InferenceClassificationSchema,
  InternetMessageHeaderSchema,
  type LeanMessage,
} from "./schemas.ts";

const GraphRecipientSchema = z.object({
  emailAddress: z.object({
    name: z.string().optional(),
    address: z.string(),
  }),
});

const GraphMessageSchema = z.object({
  id: z.string(),
  receivedDateTime: z.string(),
  hasAttachments: z.boolean(),
  subject: z.string().nullable().optional(),
  bodyPreview: z.string().nullable().optional(),
  importance: ImportanceSchema.optional(),
  inferenceClassification: InferenceClassificationSchema.optional(),
  conversationId: z.string(),
  parentFolderId: z.string().optional(),
  isRead: z.boolean(),
  // Only requested by FULL_SELECT — absent on lean search results.
  webLink: z.string().optional(),
  body: z
    .object({
      contentType: z.string(),
      content: z.string(),
    })
    .optional(),
  from: GraphRecipientSchema.nullable().optional(),
  toRecipients: z.array(GraphRecipientSchema).optional(),
  ccRecipients: z.array(GraphRecipientSchema).optional(),
  internetMessageHeaders: z.array(InternetMessageHeaderSchema).optional(),
});

const GraphFolderSchema = z.object({
  id: z.string(),
  displayName: z.string(),
  parentFolderId: z.string().nullable().optional(),
  totalItemCount: z.number().int(),
  unreadItemCount: z.number().int(),
});

type GraphMessage = z.infer<typeof GraphMessageSchema>;
type GraphRecipient = z.infer<typeof GraphRecipientSchema>;

// Invisible / formatting characters that marketing emails inject for tracking.
// U+034F combining grapheme joiner; U+200B-200F zero-width + bidi marks;
// U+202A-202E bidi overrides; U+2060-206F word joiner / invisible operators;
// U+FEFF BOM. Regular whitespace is preserved here and collapsed below.
const NOISE_CHAR_RE = /[͏​-‏‪-‮⁠-⁯﻿]/g;

export function cleanPreview(s: string): string {
  return s.replace(NOISE_CHAR_RE, "").replace(/\s+/g, " ").trim();
}

function flattenAddress(r: GraphRecipient): EmailAddress {
  const { name, address } = r.emailAddress;
  return name === undefined ? { address } : { name, address };
}

function leanFromParsed(m: GraphMessage): LeanMessage {
  return {
    id: m.id,
    from: m.from ? flattenAddress(m.from) : { address: "" },
    subject: m.subject ?? "",
    received_at: m.receivedDateTime,
    has_attachment: m.hasAttachments,
    is_read: m.isRead,
    conversation_id: m.conversationId,
    body_preview: cleanPreview(m.bodyPreview ?? ""),
    ...(m.inferenceClassification ? { inference_classification: m.inferenceClassification } : {}),
  };
}

export function mapLeanMessage(raw: unknown): LeanMessage {
  return leanFromParsed(GraphMessageSchema.parse(raw));
}

/**
 * search() needs the raw parentFolderId to resolve display names after the
 * fetch loop completes. This variant returns the lean message plus that id.
 */
export function mapLeanWithFolderId(raw: unknown): {
  lean: LeanMessage;
  folder_id: string | undefined;
} {
  const m = GraphMessageSchema.parse(raw);
  return { lean: leanFromParsed(m), folder_id: m.parentFolderId };
}

export function mapFullMessage(raw: unknown): FullMessage {
  const m = GraphMessageSchema.parse(raw);
  const contentTypeRaw = m.body?.contentType.toLowerCase();
  return {
    ...leanFromParsed(m),
    to: m.toRecipients?.map(flattenAddress) ?? [],
    cc: m.ccRecipients?.map(flattenAddress) ?? [],
    body: m.body?.content ?? "",
    body_content_type: contentTypeRaw === "html" ? "html" : "text",
    importance: m.importance ?? "normal",
    internet_message_headers: m.internetMessageHeaders ?? [],
    web_link: m.webLink ?? "",
  };
}

export function mapFolder(raw: unknown): Folder {
  const f = GraphFolderSchema.parse(raw);
  return {
    id: f.id,
    display_name: f.displayName,
    ...(f.parentFolderId ? { parent_folder_id: f.parentFolderId } : {}),
    total_item_count: f.totalItemCount,
    unread_item_count: f.unreadItemCount,
  };
}
