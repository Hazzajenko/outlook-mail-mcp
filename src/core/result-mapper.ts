import { z } from "zod";
import {
  type EmailAddress,
  type Folder,
  type FullMessage,
  ImportanceSchema,
  InferenceClassificationSchema,
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
  webLink: z.string(),
  body: z
    .object({
      contentType: z.string(),
      content: z.string(),
    })
    .optional(),
  from: GraphRecipientSchema.nullable().optional(),
  toRecipients: z.array(GraphRecipientSchema).optional(),
  ccRecipients: z.array(GraphRecipientSchema).optional(),
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
    web_link: m.webLink,
    body_preview: m.bodyPreview ?? "",
    ...(m.inferenceClassification ? { inference_classification: m.inferenceClassification } : {}),
    ...(m.parentFolderId ? { parent_folder_id: m.parentFolderId } : {}),
  };
}

export function mapLeanMessage(raw: unknown): LeanMessage {
  return leanFromParsed(GraphMessageSchema.parse(raw));
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
