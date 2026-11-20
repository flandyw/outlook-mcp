import { EventEmitter } from "node:events";
import { mkdtemp, writeFile, rm } from "node:fs/promises";
import { join } from "node:path";
import { tmpdir } from "node:os";
import { Outlook } from "../../dist/outlook.js";
import { Calendar } from "../../dist/calendar.js";

const id = "11111111-1111-4111-8111-111111111111";
const mail = { message_id: id, sender: "Ada", subject: "Fixture mail", received_at: null,
  to: null, cc: null, attachment_names: ["fixture.txt"],
  attachments: [{ index: 0, name: "fixture.txt", size_bytes: 7, type: "Text" }], body_markdown: "Fixture body", truncated: false };

Outlook.prototype.ensurePage = async function () {
  if (!this.context) {
    this.context = new EventEmitter();
    this.context.close = async () => {};
  }
  return { isClosed: () => false, url: () => "https://outlook.office.com/mail/" };
};
Outlook.prototype.ensureMailbox = async () => {};
Outlook.prototype.signInPrompt = async () => false;
Outlook.prototype.search = function (query, limit, signal) {
  if (query === "hold") console.log("fixture enqueued");
  return this.run(async () => {
    if (query === "wait" || query === "hold") {
      if (query === "wait") console.log("fixture waiting");
      await new Promise(resolve => signal.addEventListener("abort", resolve, { once: true }));
      this.checkpoint();
    }
    this.references.set(id, { query, row: {}, expires: Date.now() + 60000 });
    return { scope: "rendered_search_results", ambiguous_results_omitted: 0,
      messages: [{ message_id: id, sender: mail.sender, subject: mail.subject, received_at: null, has_attachments: true, snippet: "Fixture", unread: true }] };
  }, signal);
};
Outlook.prototype.list = function (folder, limit, unreadOnly, signal) {
  return this.run(async () => ({ scope: "rendered_folder_list", folder, scan_limit_reached: false, truncated: false,
    ambiguous_results_omitted: 0, messages: [{ message_id: id, sender: mail.sender, subject: mail.subject,
      received_at: null, has_attachments: true, snippet: "Fresh fixture", unread: true }] }), signal);
};
Outlook.prototype.readOnPage = async function (page, ref, message_id, limit) {
  return { ...mail, message_id, body_markdown: mail.body_markdown.slice(0, limit), truncated: limit < mail.body_markdown.length };
};
Outlook.prototype.downloadAttachment = async function (message_id, attachment, maxBytes, consume, signal) {
  return this.run(async () => {
    const dir = await mkdtemp(join(tmpdir(), "mcp-file-fixture-"));
    try {
      const path = join(dir, "download");
      const head = Buffer.from("fixture");
      await writeFile(path, head);
      const file = { path, head, name: "fixture.txt", size_bytes: head.length };
      return { message_id, name: file.name, size_bytes: file.size_bytes, ...await consume(file) };
    } finally { await rm(dir, { recursive: true, force: true }); }
  }, signal);
};
Calendar.prototype.search = function (input, signal) {
  return this.outlook.useBrowser(async () => {
    const event = this.remember({ date: "2026-11-12", title: "Fixture event", start: "2026-11-12", end: "2026-11-12",
      allDay: true, location: null, fingerprint: "fixture" });
    return { scope: "rendered_day_view", start: event.date, end: event.date, query: null, events: [event],
      scanned_start: event.date, scanned_end: event.date, stopped_early: false, truncated: false };
  }, signal);
};
Calendar.prototype.readInContext = async function (context, ref, event_id) {
  return { event_id, title: ref.title, start: ref.start, end: ref.end, all_day: ref.allDay, date: ref.date,
    location: null, when: null, recurrence: null, attendees: [], body_markdown: "Fixture notes", truncated: false };
};
