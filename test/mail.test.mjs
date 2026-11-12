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
  assert.match(hits.messages[0].message_id, /^ow1\./);
  assert.match(hits.messages[1].message_id, /^ow1\./);
  assert.notEqual(hits.messages[0].message_id, hits.messages[1].message_id);
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

// Outlook paints only the rows near the viewport (here, 4 of 30). Search scrolls for more, and opening
// a message jumps the list, so an earlier row may leave the DOM until the list is scrolled back to it.
test("honors limit past the first painted page and can reopen a virtualized row", async t => {
  const browser = await launchChromium();
  t.after(() => browser.close());
  const context = await browser.newContext();
  await context.route("https://outlook.office.com/**", route => route.fulfill({
    contentType: "text/html",
    body: `<input id="topSearchInput" aria-label="Search">
      <div id="MailList" hidden><div id="scroller" style="height:120px;overflow-y:auto">
        <div id="spacer" style="position:relative;height:1200px"><div role="listbox" id="rows"></div></div></div></div>
      <div role="main" aria-label="Reading Pane"><span role="heading" id="subject"></span>
        <div id="UniqueMessageBody_1" role="document" aria-label="Message body"></div></div>
      <script>
      const scroller = document.querySelector('#scroller'), list = document.querySelector('#rows');
      const render = () => {
        const first = Math.floor(scroller.scrollTop / 40), html = [];
        for (let n = first; n < Math.min(30, first + 4); n++) html.push('<div role="option" data-convid="c' + n +
          '" style="position:absolute;top:' + n * 40 + 'px;height:40px"><span data-testid="sender">Sender ' + n +
          '</span> <span data-testid="subject">Subject ' + n + '</span></div>');
        list.innerHTML = html.join('');
        for (const row of list.children) row.addEventListener('click', () => {
          const n = row.dataset.convid.slice(1);
          document.querySelector('#subject').textContent = 'Subject ' + n;
          document.querySelector('#UniqueMessageBody_1').textContent = 'Body ' + n;
          scroller.scrollTop = 1000; // Outlook moves the list after opening a message.
        });
      };
      scroller.addEventListener('scroll', render);
      document.querySelector('input').addEventListener('keydown', event => {
        if (event.key !== 'Enter') return;
        fetch('/searchservice/api/v2/query', { method: 'POST' }).then(() => {
          document.querySelector('#MailList').hidden = false;
          render();
        });
      });
      </script>`,
  }));
  const outlook = new Outlook(async () => context);
  const first = await outlook.search("subject", 1);
  assert.deepEqual(first.messages.map(message => message.subject), ["Subject 0"]);
  const hits = await outlook.search("subject", 25);
  assert.deepEqual(hits.messages.map(message => message.subject),
    Array.from({ length: 25 }, (_, index) => "Subject " + index));
  assert.match(hits.messages[24].message_id, /^ow1\./);
  assert.equal((await outlook.read(hits.messages[24].message_id)).body_markdown, "Body 24");
  // Opening a message jumps the list to the bottom, so these rows are no longer in the DOM.
  assert.equal((await outlook.read(hits.messages[3].message_id)).body_markdown, "Body 3");
  assert.equal((await outlook.read(hits.messages[1].message_id)).body_markdown, "Body 1");
  assert.equal((await outlook.read(hits.messages[0].message_id)).body_markdown, "Body 0");
});

test("rejects malformed message ids before opening the browser", async () => {
  const outlook = new Outlook(async () => { throw new Error("browser should stay closed"); });
  const incomplete = "ow1." + Buffer.from(JSON.stringify({ v: 1, q: "courier" })).toString("base64url");
  await assert.rejects(outlook.read("ow1.bm90LWpzb24"), { code: "UNKNOWN_MESSAGE" });
  await assert.rejects(outlook.read(incomplete), { code: "UNKNOWN_MESSAGE" });
  await assert.rejects(outlook.downloadAttachment(incomplete, 0, 100), { code: "UNKNOWN_MESSAGE" });
});

test("keeps stable ids usable after expiry and restart, and temporary ids until they expire", async t => {
  const browser = await launchChromium();
  t.after(() => browser.close());
  const context = await browser.newContext();
  await context.route("https://outlook.office.com/**", route => route.fulfill({
    contentType: "text/html",
    body: `<input id="topSearchInput" aria-label="Search">
      <div id="MailList" hidden><div role="listbox" id="rows"></div></div>
      <div role="main" aria-label="Reading Pane">
        <span role="heading" id="subject"></span>
        <div id="UniqueMessageBody_1" role="document" aria-label="Message body">
          <p id="body-text"></p>
          <button type="button" id="decoy" aria-label="Mark as unread" style="width:24px;height:24px"></button>
        </div>
        <button type="button" id="unread">Mark as unread</button>
      </div>
      <script>
      window.__restored = 0;
      window.__bodyClicks = 0;
      document.querySelector('#decoy').addEventListener('click', () => { window.__bodyClicks += 1; });
      document.querySelector('#unread').addEventListener('click', () => { window.__restored += 1; });
      const paint = () => {
        document.querySelector('#rows').innerHTML = \`
          <div role="option" aria-label="Unread Ada Durable note">
            <span data-item-id="item-durable" hidden></span>
            <span data-testid="sender">Ada</span>
            <span data-testid="subject">Durable note</span>
            <span data-testid="preview">Kept</span>
            <time datetime="2026-11-01T00:00:00Z"></time>
          </div>
          <div role="option" aria-label="Bea Temporary note">
            <span data-testid="sender">Bea</span>
            <span data-testid="subject">Temporary note</span>
            <span data-testid="preview">Ephemeral</span>
            <time datetime="2026-11-02T00:00:00Z"></time>
          </div>
          <div role="option" data-convid="conv-thread" aria-label="Cara Thread note">
            <span data-testid="sender">Cara</span>
            <span data-testid="subject">Thread note</span>
            <span data-testid="preview">Earlier</span>
            <time datetime="2026-11-03T00:00:00Z"></time>
          </div>
          <div role="option" data-convid="conv-thread" aria-label="Cara Thread note">
            <span data-testid="sender">Cara</span>
            <span data-testid="subject">Thread note</span>
            <span data-testid="preview">Later</span>
            <time datetime="2026-11-04T00:00:00Z"></time>
          </div>
          <div role="option" aria-label="Ada Unread books are due">
            <span data-item-id="item-subject"></span>
            <span data-testid="sender">Ada</span>
            <span data-testid="subject">Unread books are due</span>
            <span data-testid="preview">Not a status</span>
            <time datetime="2026-11-05T00:00:00Z"></time>
          </div>\`;
        for (const option of document.querySelectorAll('#rows [role=option]')) {
          option.addEventListener('click', () => {
            const subject = option.querySelector('[data-testid=subject]').textContent;
            document.querySelector('#subject').textContent = subject;
            document.querySelector('#body-text').textContent = 'Body ' + subject;
          });
        }
      };
      document.querySelector('input').addEventListener('keydown', event => {
        if (event.key !== 'Enter') return;
        fetch('/searchservice/api/v2/query', { method: 'POST' }).then(() => {
          document.querySelector('#MailList').hidden = false;
          paint();
        });
      });
      </script>`,
  }));
  const bySubject = (result, subject) => result.messages.find(message => message.subject === subject);
  const flags = () => context.pages()[0].evaluate(() => ({ restored: window.__restored, bodyClicks: window.__bodyClicks }));
  const outlook = new Outlook(async () => context);
  const hits = await outlook.search("note", 10);
  const stable = bySubject(hits, "Durable note").message_id;
  const temporary = bySubject(hits, "Temporary note").message_id;
  const earlier = bySubject(hits, "Thread note");
  const later = hits.messages.filter(message => message.subject === "Thread note")[1];
  assert.match(stable, /^ow1\./);
  assert.match(earlier.message_id, /^ow1\./);
  assert.match(later.message_id, /^ow1\./);
  assert.notEqual(earlier.message_id, later.message_id);
  assert.match(temporary, /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i);
  const again = await outlook.search("note", 10);
  assert.equal(bySubject(again, "Durable note").message_id, stable);
  assert.notEqual(bySubject(again, "Temporary note").message_id, temporary);
  assert.deepEqual(await flags(), { restored: 0, bodyClicks: 0 });
  // A subject that merely contains the word is not an unread status, so the toolbar button stays unclicked.
  assert.equal((await outlook.read(bySubject(hits, "Unread books are due").message_id)).body_markdown, "Body Unread books are due");
  assert.deepEqual(await flags(), { restored: 0, bodyClicks: 0 });
  const durable = await outlook.read(stable);
  assert.equal(durable.body_markdown, "Body Durable note");
  assert.equal(durable.received_at, "2026-11-01T00:00:00Z");
  assert.deepEqual(await flags(), { restored: 1, bodyClicks: 0 });
  // The UUID from the first search is still in flight after a later search issued a new one.
  assert.equal((await outlook.read(temporary)).subject, "Temporary note");
  assert.deepEqual(await flags(), { restored: 1, bodyClicks: 0 });

  const restarted = new Outlook(async () => context);
  assert.equal((await restarted.read(stable)).subject, "Durable note");
  assert.equal((await restarted.read(later.message_id)).received_at, "2026-11-04T00:00:00Z");
  assert.equal((await restarted.read(earlier.message_id)).received_at, "2026-11-03T00:00:00Z");
  await assert.rejects(restarted.read(temporary), { code: "UNKNOWN_MESSAGE" });
  assert.equal((await flags()).bodyClicks, 0);

  const short = new Outlook(async () => context, { referenceTtlMs: 30 });
  const shortHits = await short.search("note", 10);
  const shortStable = bySubject(shortHits, "Durable note").message_id;
  const shortTemporary = bySubject(shortHits, "Temporary note").message_id;
  assert.equal(shortStable, stable);
  await new Promise(resolve => setTimeout(resolve, 80));
  assert.equal((await short.read(shortStable)).subject, "Durable note");
  await assert.rejects(short.read(shortTemporary), { code: "UNKNOWN_MESSAGE" });
  assert.equal((await outlook.read(temporary)).subject, "Temporary note");
});
