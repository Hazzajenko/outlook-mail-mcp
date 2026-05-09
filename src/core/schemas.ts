import { z } from "zod";
import { parseDateInput } from "./date-input.ts";

const dateInput = z.string().transform((v, ctx) => {
  try {
    return parseDateInput(v);
  } catch (e) {
    ctx.addIssue({ code: "custom", message: (e as Error).message });
    return z.NEVER;
  }
});

export const ImportanceSchema = z.enum(["low", "normal", "high"]);
export type Importance = z.infer<typeof ImportanceSchema>;

export const SearchParamsSchema = z.strictObject({
  query: z.string().optional(),
  from: z.string().optional(),
  to: z.string().optional(),
  subject_contains: z.string().optional(),
  body_contains: z.string().optional(),
  since: dateInput.optional(),
  until: dateInput.optional(),
  has_attachment: z.boolean().optional(),
  folder: z.string().optional(),
  is_unread: z.boolean().optional(),
  importance: ImportanceSchema.optional(),
  top: z.number().int().min(1).max(500).default(50),
});

export type SearchParamsInput = z.input<typeof SearchParamsSchema>;
export type SearchParams = z.output<typeof SearchParamsSchema>;

const EmailAddressSchema = z.object({
  name: z.string().optional(),
  address: z.string(),
});
export type EmailAddress = z.infer<typeof EmailAddressSchema>;

export const LeanMessageSchema = z.object({
  id: z.string(),
  from: EmailAddressSchema,
  subject: z.string(),
  received_at: z.string(),
  has_attachment: z.boolean(),
  is_read: z.boolean(),
  conversation_id: z.string(),
  web_link: z.string(),
  body_preview: z.string(),
});
export type LeanMessage = z.infer<typeof LeanMessageSchema>;

export const FullMessageSchema = LeanMessageSchema.extend({
  to: z.array(EmailAddressSchema),
  cc: z.array(EmailAddressSchema),
  body: z.string(),
  body_content_type: z.enum(["text", "html"]),
  importance: ImportanceSchema,
});
export type FullMessage = z.infer<typeof FullMessageSchema>;

export const FolderSchema = z.object({
  id: z.string(),
  display_name: z.string(),
  parent_folder_id: z.string().optional(),
  total_item_count: z.number().int(),
  unread_item_count: z.number().int(),
});
export type Folder = z.infer<typeof FolderSchema>;

export const SearchResultSchema = z.object({
  results: z.array(LeanMessageSchema),
  total_returned: z.number().int(),
  more_available: z.boolean(),
});
export type SearchResult = z.infer<typeof SearchResultSchema>;
