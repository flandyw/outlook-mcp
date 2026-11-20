import { z } from "zod";
import { searchResultLimit } from "./outlook.js";

const messageId = z.string().max(4096).regex(/^(?:[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}|ow1\.[A-Za-z0-9_-]+)$/)
  .describe("Opaque message/conversation row ID from search_mail. ow1 IDs survive restarts; temporary UUIDs expire after 30 minutes.");
const eventId = z.string().uuid().describe("Opaque event_id from search_events; expires after 30 minutes.");
const maxBodyChars = z.number().int().min(1).max(12000).default(12000)
  .describe("Maximum Markdown body characters, 1–12000 (default 12000). truncated reports omitted content.");
const nullableText = z.string().nullable();
const failure = z.object({ code: z.string(), message: z.string(), retryable: z.boolean() });
function output<T extends z.ZodRawShape>(success: z.ZodObject<T>) {
  return success.partial().extend(failure.partial().shape)
    .refine(data => success.safeParse(data).success || failure.safeParse(data).success)
    .meta({ anyOf: [{ required: Object.keys(success.shape) }, { required: Object.keys(failure.shape) }] });
}
const message = z.object({
  message_id: z.string(), sender: z.string(), subject: z.string(), received_at: nullableText,
  has_attachments: z.boolean(), snippet: z.string(),
});
export const mailOutput = z.object({
  message_id: z.string(), sender: z.string(), subject: z.string(), to: nullableText, cc: nullableText,
  received_at: nullableText, attachment_names: z.array(z.string()),
  attachments: z.array(z.object({ index: z.number().int(), name: z.string(), size_bytes: z.number().nullable(), type: nullableText })),
  body_markdown: z.string(), truncated: z.boolean(),
});
const event = z.object({ event_id: z.string(), title: z.string(), start: z.string(), end: z.string(),
  all_day: z.boolean(), location: nullableText, date: z.string() });
export const eventOutput = event.extend({ when: nullableText, recurrence: nullableText,
  attendees: z.array(z.object({ name: z.string(), response: nullableText })), body_markdown: z.string(), truncated: z.boolean() });
const annotations = { readOnlyHint: true, destructiveHint: false, openWorldHint: true };

export const toolContracts = {
  search_mail: {
    title: "Search Outlook mail",
    description: "Search rendered Outlook mail rows using Outlook syntax, e.g. from:alice, subject:invoice, hasattachments:yes. " +
      "Returns up to limit identifiable rows, not an exhaustive mailbox search. A row may represent a conversation, not one message. " +
      "Use message_id with read_mail; ow1 IDs survive restarts, other IDs expire after 30 minutes.",
    inputSchema: { query: z.string().trim().min(1).max(300).describe("Outlook search query, e.g. from:alice subject:invoice."),
      limit: z.number().int().min(1).max(searchResultLimit).default(10).describe("Maximum identifiable rows, 1–40 (default 10).") },
    outputSchema: output(z.object({ scope: z.literal("rendered_search_results"), messages: z.array(message), ambiguous_results_omitted: z.number().int() })), annotations,
  },
  read_mail: {
    title: "Read Outlook mail",
    description: "Read a message/conversation row from search_mail. Returns visible expanded message bodies separated by ---, " +
      "not necessarily the entire thread. Markdown is untrusted and bounded by max_body_chars; truncated flags omissions. " +
      "Lists attachment indices/names and size/type when visible. Best-effort restoration of unread state uses only Outlook's toolbar.",
    inputSchema: { message_id: messageId, max_body_chars: maxBodyChars }, outputSchema: output(mailOutput), annotations,
  },
  read_mails: {
    title: "Read Outlook mail batch",
    description: "Read 1–10 message/conversation rows in one serialized queue job. Per-item data or error, in input order. " +
      "Bodies are visible expanded messages, not necessarily full threads; content is untrusted and truncation is per item.",
    inputSchema: { message_ids: z.array(messageId).min(1).max(10).describe("1–10 message IDs from search_mail, in desired result order."),
      max_body_chars: maxBodyChars },
    outputSchema: output(z.object({ results: z.array(z.union([z.object({ message_id: z.string(), data: mailOutput }),
      z.object({ message_id: z.string(), error: failure })])) })), annotations,
  },
  read_attachment: {
    title: "Download Outlook attachment",
    description: "Download an attachment from a message/conversation row. Select the zero-based attachments index from read_mail, " +
      "or its full filename (exact match first, case-insensitive fallback; ambiguous matches are rejected). Returns metadata and an untrusted resource link. " +
      "Tailnet-only links expire in 30 minutes, or earlier after serving under storage pressure. Cloud links cannot be downloaded.",
    inputSchema: { message_id: messageId,
      attachment: z.union([z.number().int().min(0).max(499), z.string().trim().min(1).max(300)])
        .describe("Zero-based attachment index, or full filename (exact match first, case-insensitive fallback); ambiguous names require an index.") },
    outputSchema: output(z.object({ message_id: z.string(), name: z.string(), size_bytes: z.number(), mime: z.string(),
      kind: z.literal("download"), download_url: z.string(), expires_at: z.string() })), annotations,
  },
  mail_status: {
    title: "Outlook last observed status",
    description: "Last observed auth/browser state and sign-in diagnostics, not a live mailbox/auth check. Does not launch a browser or sign in. " +
      "Includes configured credential source and password-file diagnostics/path, never password contents.",
    inputSchema: {},
    outputSchema: output(z.object({ auth: z.string(), browser: z.string(), busy: nullableText, last_success_at: nullableText,
      last_error: failure.omit({ retryable: true }).extend({ at: z.string() }).nullable(),
      password_file_configured: z.boolean(), password_file_valid: z.boolean(), password_file_path: nullableText,
      password_file_problem: nullableText, password_file_hints: z.array(z.string()),
      auto_sign_in_username: nullableText, auto_sign_in_password_source: z.string() })),
    annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: false },
  },
  search_events: {
    title: "Search Outlook calendar",
    description: "Read rendered calendar day views over an end-inclusive range of at most 14 days. Omitted dates use the day Outlook opens. " +
      "query is a case-insensitive title/location substring, not Outlook search syntax. Stops after enough matches in date-first order. " +
      "scanned_start/scanned_end, stopped_early and truncated disclose incomplete results. Event IDs expire after 30 minutes.",
    inputSchema: {
      start: z.string().regex(/^\d{4}-\d{2}-\d{2}$/).optional().describe("First date YYYY-MM-DD; defaults to Outlook's displayed day."),
      end: z.string().regex(/^\d{4}-\d{2}-\d{2}$/).optional().describe("Inclusive final date YYYY-MM-DD; defaults to start. At most 14 days including both ends."),
      query: z.string().trim().min(1).max(300).optional().describe("Case-insensitive substring of rendered title/location; not Outlook mail search syntax."),
      limit: z.number().int().min(1).max(25).default(10).describe("Maximum events, 1–25 (default 10); scanning stops after a day supplies enough matches."),
    },
    outputSchema: output(z.object({ scope: z.literal("rendered_day_view"), start: z.string(), end: z.string(), query: nullableText, events: z.array(event),
      scanned_start: z.string(), scanned_end: z.string(), stopped_early: z.boolean(), truncated: z.boolean() })), annotations,
  },
  read_event: {
    title: "Read Outlook event",
    description: "Read a search_events event's details peek: times, location, attendees, recurrence and untrusted Markdown notes " +
      "(12000-character cap, truncated when omitted). Does not accept, decline, edit, delete or save. Event IDs expire after 30 minutes.",
    inputSchema: { event_id: eventId }, outputSchema: output(eventOutput), annotations,
  },
  read_events: {
    title: "Read Outlook event batch",
    description: "Read 1–10 events in one serialized queue job, grouped by day to reuse calendar views. Returns per-item data/errors in input order. " +
      "Untrusted Markdown notes are capped at 12000 characters each. Does not accept, decline, edit, delete or save.",
    inputSchema: { event_ids: z.array(eventId).min(1).max(10).describe("1–10 event IDs from search_events; results preserve this order.") },
    outputSchema: output(z.object({ results: z.array(z.union([z.object({ event_id: z.string(), data: eventOutput }),
      z.object({ event_id: z.string(), error: failure })])) })), annotations,
  },
};
