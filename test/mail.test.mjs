import { test } from "node:test";
import assert from "node:assert/strict";
import { launchChromium } from "./chromium.mjs";
import { Outlook } from "../dist/outlook.js";

test("searches and reads a message using the browser", async t => {
  const browser = await launchChromium();
  t.after(() => browser.close());
  const context = await browser.newContext();
  await context.route("https://outlook.office.com/**", route => route.fulfill({
    contentType: "text/html",
    body: `<input type="search" aria-label="Search">
      <section aria-label="Search results" hidden>
        <div role="listbox"><div role="option" data-item-id="1">
          <span data-testid="sender">Maya Chen</span>
          <span data-testid="subject">Delivery window</span>
          <span data-testid="preview">The courier confirmed Friday.</span>
          <time datetime="2026-11-11T10:42:00+11:00">10:42</time>
        </div></div>
      </section>
      <section aria-label="Reading Pane" hidden>
        <h2>Delivery window</h2>
        <div role="document"><p>The courier confirmed Friday.</p>
          <ul><li>Clear the loading bay</li></ul>
          <p hidden>Hidden content</p>
          <p><a href="https://example.com/private">Delivery details</a></p>
        </div>
        <span data-testid="attachment-name">delivery.pdf</span>
      </section>
      <script>
      document.querySelector('input').addEventListener('keydown', event => {
        if (event.key === 'Enter') document.querySelector('section').hidden = false;
      });
      document.querySelector('[role="option"]').addEventListener('click', () => {
        document.querySelector('[aria-label="Reading Pane"]').hidden = false;
      });
      </script>`,
  }));
  const outlook = new Outlook(async () => context);
  const hits = await outlook.search("courier", 10);
  assert.equal(hits.messages.length, 1);
  const message = await outlook.read(hits.messages[0].message_id);
  assert.equal(message.subject, "Delivery window");
  assert.equal(message.received_at, "2026-11-11T10:42:00+11:00");
  assert.deepEqual(message.attachment_names, ["delivery.pdf"]);
  assert.equal(message.body_markdown, "The courier confirmed Friday.\n\n- Clear the loading bay\n\nDelivery details");
  await assert.rejects(outlook.read("unknown"), { code: "UNKNOWN_MESSAGE" });
});

// Mirrors the structure observed in the current Outlook Web layout (outlook.cloud.microsoft, Nov 2026).
test("reads the real body, recipients and attachments from the modern layout", async t => {
  const browser = await launchChromium();
  t.after(() => browser.close());
  const context = await browser.newContext();
  const row = (time, preview) => `<div role="option" data-convid="conv-1" aria-label="Maya Chen Quarterly plan Has attachments">
      <div class="ESO13"><span title="maya@example.com">Maya Chen</span></div>
      <span class="TtcXM">Quarterly plan</span><span class="ASFJj">${preview}</span>
      <span class="qq2gS" title="${time}">${time}</span></div>`;
  await context.route("https://outlook.office.com/**", route => route.fulfill({
    contentType: "text/html",
    body: `<input id="topSearchInput" aria-label="Search">
      <div id="MailList" hidden><div role="listbox">
        ${row("Wed 2026-11-11 09:00", "Draft attached")}
        ${row("Tue 2026-11-10 08:00", "Earlier note")}
        ${row("Wed 2026-11-11 09:00", "Draft attached")}
      </div></div>
      <div role="main" aria-label="Reading Pane" hidden>
        <div class="allowTextSelection"><span id="CONV_1_SUBJECT" role="heading">Quarterly plan</span></div>
        <div id="MSG_1_TO" role="heading">To: Sam Lee; Ana Ruiz</div>
        <div class="allowTextSelection">
          <div id="UniqueMessageBody_1" role="document" aria-label="Message body">
            <div>Hi <b>Sam</b>, the plan is <a href="https://example.com">here</a>.</div>
            <table><tr><td>Owner</td><td>Due</td></tr><tr><td>Maya</td><td>Friday</td></tr></table>
            <div>Line one<br>Line two</div>
            <img alt="Company logo" src="logo.png">
          </div>
        </div>
        <div role="listbox" aria-label="file attachments">
          <div role="option" aria-label="plan.pdf Open 3 MB"><div title="plan.pdf">plan.pdf</div><div>3 MB</div>
            <button aria-label="More actions" aria-haspopup="true" data-file="plan.pdf" data-body="Plan body">⋯</button></div>
          <div role="option" aria-label="notes.txt Open 1 KB"><div title="notes.txt">notes.txt</div><div>1 KB</div>
            <button aria-label="More actions" aria-haspopup="true" data-file="notes.txt" data-body="Remember the milk">⋯</button></div>
          <div role="option" aria-label="Shared link Open"><div title="Shared link">Shared link</div>
            <button aria-label="More actions" aria-haspopup="true">⋯</button></div>
        </div>
        <div role="listbox" aria-label="image attachments">
          <div role="option" aria-label="photo.jpg Open 2 MB"><img alt="">
            <button aria-label="More actions" aria-haspopup="true" data-file="photo.jpg" data-body="JPEG bytes"></button></div>
        </div>
        <div role="menu" hidden><div role="menuitem">Preview</div><div role="menuitem" id="download">Download</div></div>
      </div>
      <script>
      document.querySelector('input').addEventListener('keydown', event => {
        if (event.key !== 'Enter') return;
        fetch('/searchservice/api/v2/query', { method: 'POST' })
          .then(() => { document.querySelector('#MailList').hidden = false; });
      });
      let pending;
      for (const button of document.querySelectorAll('[aria-haspopup]')) button.addEventListener('click', () => {
        pending = button;
        const menu = document.querySelector('[role=menu]');
        menu.hidden = false;
        document.querySelector('#download').hidden = !button.dataset.file; // cloud links have no Download
      });
      document.querySelector('#download').addEventListener('click', () => {
        document.querySelector('[role=menu]').hidden = true;
        const link = document.createElement('a');
        link.download = pending.dataset.file;
        link.href = URL.createObjectURL(new Blob([pending.dataset.body]));
        link.click();
      });
      for (const option of document.querySelectorAll('#MailList [role="option"]')) option.addEventListener('click', () => {
        document.querySelector('[aria-label="Reading Pane"]').hidden = false;
        // Outlook re-renders previews once a message is opened.
        for (const preview of document.querySelectorAll('.ASFJj')) preview.textContent += " (seen)";
      });
      </script>`,
  }));
  const outlook = new Outlook(async () => context);
  const hits = await outlook.search("plan", 10);
  assert.equal(hits.messages.length, 2, "duplicate keyed rows collapse to one result");
  assert.equal(hits.ambiguous_results_omitted, 0);
  assert.equal(hits.messages[0].has_attachments, true);
  const message = await outlook.read(hits.messages[0].message_id);
  assert.equal(message.body_markdown,
    "Hi Sam, the plan is here.\n\nOwner | Due\n\nMaya | Friday\n\nLine one\nLine two\n\n(image: Company logo)");
  assert.equal(message.to, "Sam Lee; Ana Ruiz");
  assert.equal(message.cc, null);
  assert.deepEqual(message.attachment_names, ["plan.pdf", "notes.txt", "Shared link", "photo.jpg"]);
  const id = hits.messages[0].message_id;
  const byIndex = await outlook.downloadAttachment(id, 1, 1024);
  assert.deepEqual([byIndex.name, byIndex.size_bytes, byIndex.data.toString()], ["notes.txt", 17, "Remember the milk"]);
  const byName = await outlook.downloadAttachment(id, "PLAN.PDF", 1024);
  assert.equal(byName.data.toString(), "Plan body");
  await assert.rejects(outlook.downloadAttachment(id, "missing.doc", 1024), { code: "UNKNOWN_ATTACHMENT", message: /notes\.txt/ });
  await assert.rejects(outlook.downloadAttachment(id, 7, 1024), { code: "UNKNOWN_ATTACHMENT" });
  await assert.rejects(outlook.downloadAttachment(id, 2, 1024), { code: "ATTACHMENT_UNAVAILABLE" });
  const photo = await outlook.downloadAttachment(id, 3, 1024);
  assert.deepEqual([photo.name, photo.data.toString()], ["photo.jpg", "JPEG bytes"]);
  await assert.rejects(outlook.downloadAttachment(id, 0, 4), { code: "ATTACHMENT_TOO_LARGE" });
  // The second message shares the conversation key, sender and subject. After previews change it must still resolve.
  const second = await outlook.read(hits.messages[1].message_id);
  assert.equal(second.received_at, "Tue 2026-11-10 08:00");
});
