import { test } from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { launchChromium } from "./chromium.mjs";
import { Outlook, mailFolders } from "../dist/outlook.js";
import { toolContracts } from "../dist/tool-contracts.js";

test("lists live folders without search and reopens virtualized folder IDs across views and restarts", async t => {
  const browser = await launchChromium();
  t.after(() => browser.close());
  const context = await browser.newContext();
  const navigations = [];
  let bounded = false;
  let clearSearch = true;
  await context.route("https://outlook.office.com/**", route => {
    const path = new URL(route.request().url()).pathname;
    navigations.push(path);
    const folder = path.split("/")[2];
    const empty = folder === "archive";
    return route.fulfill({ contentType: "text/html", body: `
      <input type="search" value="${folder && clearSearch ? "" : "stale search"}" aria-label="Search">
      <div id="MailList"><div id="scroller" style="height:100px;overflow:auto;position:relative">
      <div id="rows" role="listbox" style="height:${empty ? 0 : bounded ? 8000 : 320}px"></div></div></div>
      ${empty ? '<p>This folder is empty</p>' : ''}
      <section aria-label="Reading Pane" hidden><h2></h2><div role="document"></div>
      <span data-testid="attachment-name">fresh.txt<button aria-haspopup="true">More</button></span>
      <div role="toolbar"><button aria-label="Mark as unread">Mark as unread</button></div></section>
      <div role="menu" hidden><button role="menuitem">Download</button></div>
      <script>
      window.restored=0; window.opened=0; window.searches=0;
      document.querySelector('input').onkeydown = e => { if(e.key==='Enter') window.searches++; };
      const scroller=document.querySelector('#scroller'), rows=document.querySelector('#rows');
      function render() {
        rows.innerHTML='';
        const first=Math.floor(scroller.scrollTop/40);
        for(let n=first;n<Math.min(${empty ? 0 : 8},first+3);n++) {
          const i=n>=6?0:n;
          rows.insertAdjacentHTML('beforeend','<div role="option" '+(n<6?'data-item-id="fresh-'+n+'"':'')+
            ' aria-label="'+(n%2?'Unread ':'Read ')+'Has attachments" style="position:absolute;top:'+n*40+'px;height:40px">'+
            '<span data-testid="sender">Ada</span><span data-testid="subject">Fresh '+i+'</span>'+
            '<span data-testid="preview">New arrival '+i+'</span><time datetime="2026-11-'+(10+i)+'T10:00:00Z"></time></div>');
        }
        for(const row of rows.children) row.onclick=()=>{
          window.opened++; document.querySelector('section').hidden=false;
          document.querySelector('h2').textContent=row.querySelector('[data-testid=subject]').textContent;
          document.querySelector('[role=document]').textContent='Live body';
        };
      }
      scroller.onscroll=()=>setTimeout(render,10); render();
      document.querySelector('[aria-label="Mark as unread"]').onclick=()=>window.restored++;
      document.querySelector('[aria-haspopup]').onclick=()=>document.querySelector('[role=menu]').hidden=false;
      document.querySelector('[role=menuitem]').onclick=()=>{
        document.querySelector('[role=menu]').hidden=true;
        const a=document.createElement('a');a.download='fresh.txt';a.href=URL.createObjectURL(new Blob(['fresh']));a.click();
      };
      </script>` });
  });
  const outlook = new Outlook(async () => context);
  const listed = await outlook.list();
  toolContracts.list_mail.outputSchema.parse(listed);
  assert.equal(listed.folder, "inbox");
  assert.deepEqual(listed.messages.map(m => m.subject), [5, 4, 3, 2, 1, 0].map(n => "Fresh " + n));
  assert.equal(listed.ambiguous_results_omitted, 2);
  assert.equal(listed.scan_limit_reached, false);
  const page = context.pages()[0];
  assert.equal(await page.locator('input').inputValue(), "");
  assert.equal(await page.evaluate(() => window.searches), 0);
  const unread = await outlook.list("inbox", 2, true);
  assert.deepEqual(unread.messages.map(m => m.subject), ["Fresh 5", "Fresh 3"]);
  assert.ok(unread.messages.every(m => m.unread));
  assert.equal(unread.truncated, true);
  await page.goto("https://outlook.office.com/mail/archive/");
  const restarted = new Outlook(async () => context);
  const id = listed.messages[0].message_id;
  assert.equal((await restarted.read(id)).body_markdown, "Live body");
  assert.equal(await page.evaluate(() => window.restored), 1);
  assert.equal((await restarted.readMails([id])).results[0].data.subject, "Fresh 5");
  const attachment = await restarted.downloadAttachment(id, 0, 1024, async file => ({ text: await readFile(file.path, "utf8") }));
  assert.equal(attachment.text, "fresh");
  assert.equal(await page.evaluate(() => window.searches), 0);
  for (const folder of mailFolders) {
    const result = await outlook.list(folder, 1);
    assert.equal(result.folder, folder);
    assert.ok(navigations.includes('/mail/' + folder + '/'));
    if (folder === "archive") assert.deepEqual(result.messages, []);
    if (folder === "drafts") {
      await assert.rejects(outlook.read(result.messages[0].message_id), { code: "DRAFT_UNAVAILABLE" });
      assert.equal(await page.evaluate(() => window.opened), 0);
    }
  }
  bounded = true;
  assert.equal((await outlook.list()).scan_limit_reached, true);
  clearSearch = false;
  await assert.rejects(outlook.list(), { code: "UI_CHANGED" });
  for (const args of [["../calendar", 1], ["inbox", 0], ["inbox", 41]]) assert.throws(() => outlook.list(...args), { code: "INVALID_RANGE" });
});

test("re-finds rows from the final listing window for reads, batches and attachments after navigation", async t => {
  const browser = await launchChromium();
  t.after(() => browser.close());
  const context = await browser.newContext();
  await context.route("https://outlook.office.com/**", route => route.fulfill({ contentType: "text/html", body: `
    <input type="search" aria-label="Search">
    <div id="MailList"><div id="scroll" style="height:100px;overflow:auto;position:relative">
      <div id="rows" role="listbox" style="height:4200px"></div></div></div>
    <section aria-label="Reading Pane" hidden><h2></h2><div role="document">Deep body</div>
      <div role="listbox" aria-label="file attachments"><div role="option" aria-label="deep.txt Open">
        <span title="deep.txt">deep.txt</span><button aria-haspopup="true">More</button></div></div></section>
    <div role="menu" hidden><button role="menuitem">Download</button></div>
    <script>
    const scroll=document.querySelector('#scroll'), rows=document.querySelector('#rows');
    function render() {
      const n=Math.floor(scroll.scrollTop/80);
      rows.innerHTML='<div role="option" data-item-id="deep-'+n+'" style="position:absolute;top:'+n*80+'px;height:100px">'+
        '<span data-testid="sender">Ada</span><span data-testid="subject">Deep '+n+'</span>'+
        '<time datetime="'+new Date(Date.UTC(2026,10,1)+n*60000).toISOString()+'"></time></div>';
      rows.firstElementChild.onclick=()=>{
        document.querySelector('section').hidden=false;
        document.querySelector('h2').textContent='Deep '+n;
      };
    }
    scroll.onscroll=render; render();
    document.querySelector('[aria-haspopup]').onclick=()=>document.querySelector('[role=menu]').hidden=false;
    document.querySelector('[role=menuitem]').onclick=()=>{
      const a=document.createElement('a'); a.download='deep.txt'; a.href=URL.createObjectURL(new Blob(['deep'])); a.click();
      document.querySelector('[role=menu]').hidden=true;
    };
    </script>` }));
  const outlook = new Outlook(async () => context);
  const listed = await outlook.list("inbox", 1);
  assert.equal(listed.messages[0].subject, "Deep 50", "the last of 50 scroll advances is selectable");
  assert.equal(listed.scan_limit_reached, true);
  const id = listed.messages[0].message_id;
  const page = context.pages()[0];
  const restarted = new Outlook(async () => context);
  await page.goto("https://outlook.office.com/mail/archive/");
  assert.equal((await restarted.read(id)).subject, "Deep 50");
  await page.goto("https://outlook.office.com/mail/archive/");
  assert.equal((await restarted.readMails([id])).results[0].data.body_markdown, "Deep body");
  await page.goto("https://outlook.office.com/mail/archive/");
  const attachment = await restarted.downloadAttachment(id, 0, 1024, async file => ({ text: await readFile(file.path, "utf8") }));
  assert.equal(attachment.text, "deep");
});

test("mixed dates sort dated slots consistently while preserving unknown positions and equal-date order", async t => {
  const browser = await launchChromium();
  t.after(() => browser.close());
  const context = await browser.newContext();
  let dates = ["2026-11-10", "Yesterday", "2026-11-12", "", "2026-11-11", "Today", "2026-11-12"];
  await context.route("https://outlook.office.com/**", route => route.fulfill({ contentType: "text/html", body: `
    <input type="search" aria-label="Search"><div id="MailList"><div role="listbox">
    ${dates.map((date, index) => `<div role="option" data-item-id="row-${index}">
      <span data-testid="sender">Ada</span><span data-testid="subject">Row ${index}</span>
      <time datetime="${date}"></time></div>`).join("")}</div></div>` }));
  const outlook = new Outlook(async () => context);
  const subjects = async limit => (await outlook.list("inbox", limit)).messages.map(row => row.subject);
  assert.deepEqual(await subjects(10), [2, 1, 6, 3, 4, 5, 0].map(index => "Row " + index));
  assert.deepEqual(await subjects(2), ["Row 2", "Row 1"], "apply limit after consistently ordering the dated slots");
  dates = ["Today", "2026-11-10", "Yesterday", "2026-11-12"];
  assert.deepEqual(await subjects(2), ["Row 0", "Row 3"], "a fresh relative-date row keeps its leading position");
  dates = ["Today", "Yesterday", ""];
  assert.deepEqual(await subjects(10), ["Row 0", "Row 1", "Row 2"]);
});
