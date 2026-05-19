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

export const InferenceClassificationSchema = z.enum(["focused", "other"]);
export type InferenceClassification = z.infer<typeof InferenceClassificationSchema>;

const filterShape = {
  query: z.string().optional().describe("KQL free-text search across subject/body/people"),
  from: z.string().optional().describe("sender address or domain (text search)"),
  to: z.string().optional().describe("recipient address or domain (text search)"),
  subject_contains: z.string().optional().describe("substring in subject (text search)"),
  body_contains: z.string().optional().describe("substring in body (text search)"),
  since: dateInput
    .optional()
    .describe(
      "lower bound, inclusive. ISO date (2026-05-08) or relative (-7d, -2w, -3h, -30m). Date-only ⇒ start of that day UTC.",
    ),
  until: dateInput
    .optional()
    .describe(
      "upper bound. Date-only (2026-05-08) ⇒ inclusive of the entire day. With explicit time ⇒ inclusive of that instant.",
    ),
  has_attachment: z.boolean().optional(),
  folder: z
    .string()
    .optional()
    .describe(
      "well-known (inbox/sentitems/junkemail/deleteditems/drafts/archive) or custom display name. When omitted, Graph $search excludes Junk/Deleted/Drafts — pass folder explicitly to search them.",
    ),
  is_unread: z.boolean().optional(),
  importance: ImportanceSchema.optional(),
  inference_classification: InferenceClassificationSchema.optional().describe(
    "focused | other. Cannot combine with text search (query/from/to/subject_contains/body_contains).",
  ),
} as const;

export const FilterParamsSchema = z.strictObject(filterShape);
export type FilterParams = z.output<typeof FilterParamsSchema>;

export const SearchParamsSchema = z.strictObject({
  ...filterShape,
  top: z
    .number()
    .int()
    .min(1)
    .max(500)
    .default(50)
    .describe("max results per call (1-500, default 50). Check has_more / next_cursor for more."),
  cursor: z.string().optional().describe("pagination cursor from a previous result's next_cursor"),
});

export type SearchParamsInput = z.input<typeof SearchParamsSchema>;
export type SearchParams = z.output<typeof SearchParamsSchema>;

export const CountParamsSchema = z.strictObject(filterShape);
export type CountParams = z.output<typeof CountParamsSchema>;

export const CountResultSchema = z.object({
  count: z.number().int(),
});
export type CountResult = z.infer<typeof CountResultSchema>;

const EmailAddressSchema = z.object({
  name: z.string().optional(),
  address: z.string(),
});
export type EmailAddress = z.infer<typeof EmailAddressSchema>;

export const InternetMessageHeaderSchema = z.object({
  name: z.string(),
  value: z.string(),
});
export type InternetMessageHeader = z.infer<typeof InternetMessageHeaderSchema>;

export const LeanMessageSchema = z.object({
  id: z.string(),
  from: EmailAddressSchema,
  subject: z.string(),
  received_at: z.string(),
  has_attachment: z.boolean(),
  is_read: z.boolean(),
  conversation_id: z.string(),
  body_preview: z.string(),
  inference_classification: InferenceClassificationSchema.optional(),
  folder: z.string().optional(),
});
export type LeanMessage = z.infer<typeof LeanMessageSchema>;

export const FullMessageSchema = LeanMessageSchema.extend({
  to: z.array(EmailAddressSchema),
  cc: z.array(EmailAddressSchema),
  body: z.string(),
  body_content_type: z.enum(["text", "html"]),
  importance: ImportanceSchema,
  internet_message_headers: z.array(InternetMessageHeaderSchema),
  web_link: z.string(),
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
  has_more: z.boolean(),
  next_cursor: z.string().optional(),
});
export type SearchResult = z.infer<typeof SearchResultSchema>;
