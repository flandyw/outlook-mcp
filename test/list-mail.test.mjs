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
