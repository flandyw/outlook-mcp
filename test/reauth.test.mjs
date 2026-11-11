import { test } from "node:test";
import { strict as assert } from "node:assert";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { launchChromium } from "./chromium.mjs";
import { reauthenticate } from "../dist/reauth.js";

test("submits a test password and waits for Outlook", async t => {
  const dir = await mkdtemp(join(tmpdir(), "outlook-reauth-"));
  t.after(() => rm(dir, { recursive: true, force: true }));
  const file = join(dir, "password");
  await writeFile(file, "fixture-password\n", { mode: 0o600 });
  const browser = await launchChromium();
  t.after(() => browser.close());
  const page = await browser.newPage();
  await page.route("https://login.microsoftonline.com/**", route => route.fulfill({
    contentType: "text/html",
    body: `<form><p>fixture@example.com</p><input type="password" name="passwd"><input type="submit" id="idSIButton9"></form>
      <script>document.forms[0].onsubmit = (event) => {
        event.preventDefault();
        if (document.querySelector('input[name=passwd]').value === 'fixture-password')
          location.href = 'https://outlook.office.com/mail/';
      };</script>`,
  }));
  await page.route("https://outlook.office.com/**", route => route.fulfill({ body: "Mailbox fixture" }));
  await page.goto("https://login.microsoftonline.com/");
  await reauthenticate(page, "fixture@example.com", file);
  assert.equal(page.url(), "https://outlook.office.com/mail/");
});
