import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, mkdir, rm, writeFile, readdir } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { launchChromium } from "./chromium.mjs";
import { Outlook } from "../dist/outlook.js";

const mailboxPage = `<input id="topSearchInput" aria-label="Search">
  <div id="MailList"><div role="listbox"></div></div>
  <script>document.querySelector('input').addEventListener('keydown', event => {
    if (event.key === 'Enter') fetch('/searchservice/api/v2/query', { method: 'POST' })
      .then(() => { document.querySelector('[role=listbox]').innerHTML = '<p>No results</p>'; });
  });</script>`;

// Playwright cannot fulfil a cross-origin 302 for a navigation, so fixtures redirect from script.
const redirect = url => ({ contentType: "text/html", body: `<script>location.replace(${JSON.stringify(url)})</script>` });

// A browser whose contexts are served by the given route handler, plus a launcher counting launches.
async function harness(t, handler) {
  const browser = await launchChromium();
  t.after(() => browser.close());
  const launcher = async () => {
    launcher.count++;
    const context = await browser.newContext();
    await context.route(/^https:\/\/(outlook\.office\.com|login\.microsoftonline\.com)\//, handler);
    launcher.contexts.push(context);
    return context;
  };
  launcher.count = 0;
  launcher.contexts = [];
  return launcher;
}

test("reports an expired session and notifies once", async t => {
  const launcher = await harness(t, route => route.request().url().startsWith("https://outlook.office.com/")
    ? route.fulfill(redirect("https://login.microsoftonline.com/common/oauth2/authorize"))
    : route.fulfill({ contentType: "text/html", body: '<p>Enter password</p><input type="password" name="passwd"><input type="submit" id="idSIButton9">' }));
  const changes = [];
  const outlook = new Outlook(launcher, { onAuthChange: state => changes.push(state), readyMs: 1000 });
  delete process.env.OUTLOOK_PASSWORD_FILE;
  await assert.rejects(outlook.search("x", 5), { code: "AUTH_REQUIRED" });
  await assert.rejects(outlook.search("x", 5), { code: "AUTH_REQUIRED" });
  assert.deepEqual(changes, ["signed_out"]);
  assert.equal(outlook.status().auth, "signed_out");
  assert.equal(outlook.status().last_error.code, "AUTH_REQUIRED");
  await outlook.close();
});

test("resumes single sign-on through the account picker and stay-signed-in prompt", async t => {
  const launcher = await harness(t, route => {
    const url = new URL(route.request().url());
    if (url.hostname === "outlook.office.com") {
      return route.fulfill(url.searchParams.has("signedin") || url.pathname.startsWith("/searchservice")
        ? { contentType: "text/html", body: mailboxPage }
        : redirect("https://login.microsoftonline.com/picker"));
    }
    return route.fulfill({ contentType: "text/html", body: url.pathname === "/picker"
      ? `<div role="listitem"><div role="button" data-test-id="sam@example.com" onclick="location.href='/kmsi'">Sam</div></div>`
      : `<div>Stay signed in?</div><input type="submit" id="idSIButton9" value="Yes"
          onclick="location.href='https://outlook.office.com/mail/?signedin'">` });
  });
  const changes = [];
  const outlook = new Outlook(launcher, { onAuthChange: state => changes.push(state), readyMs: 3000 });
  const result = await outlook.search("nothing", 5);
  assert.deepEqual(result.messages, []);
  assert.deepEqual(changes, ["signed_in"]);
  await outlook.close();
});

test("reloads once when Outlook shows an error page", async t => {
  let loads = 0;
  const launcher = await harness(t, route => {
    if (route.request().isNavigationRequest()) loads++;
    return route.fulfill({ contentType: "text/html", body: loads === 1 && route.request().isNavigationRequest()
      ? "<p>Something went wrong</p>" : mailboxPage });
  });
  const outlook = new Outlook(launcher, { readyMs: 1000 });
  assert.deepEqual((await outlook.search("nothing", 5)).messages, []);
  assert.equal(loads, 2);
  await outlook.close();
});

test("recovers from a closed browser, resets a hung one and closes when idle", async t => {
  const launcher = await harness(t, route => route.request().url().includes("/searchservice/")
    ? new Promise(() => {}) // Outlook never answers the search.
    : route.fulfill({ contentType: "text/html", body: mailboxPage }));
  const outlook = new Outlook(launcher, { operationMs: 2500, idleMs: 400, readyMs: 2000 });
  await outlook.check();
  await launcher.contexts[0].close(); // Browser crash or user closed the window.
  await outlook.check();
  assert.equal(launcher.count, 2);
  await assert.rejects(outlook.search("hang", 5), { code: "TIMEOUT" });
  assert.equal(outlook.status().browser, "stopped", "a hung browser is discarded");
  await outlook.check();
  assert.equal(launcher.count, 3);
  await new Promise(resolve => setTimeout(resolve, 800));
  assert.equal(outlook.status().browser, "stopped", "an idle browser closes");
  await outlook.close();
});

test("interactive sign-in blocks mail requests and finishes when Outlook loads", async t => {
  const visible = await harness(t, route => route.fulfill({ contentType: "text/html", body: mailboxPage }));
  const outlook = new Outlook(visible, { launchVisible: visible });
  const pending = outlook.signIn({ display: { WAYLAND_DISPLAY: "wayland-test" } });
  await new Promise(resolve => setTimeout(resolve, 100));
  await assert.rejects(outlook.search("x", 5), { code: "AUTH_REQUIRED", message: /in progress/ });
  await assert.rejects(outlook.signIn(), { code: "BUSY" });
  assert.equal(await pending, "signed_in");
  assert.equal(outlook.status().auth, "signed_in");
  assert.equal(outlook.status().busy, null);
});

test("interactive sign-in reports a closed window", async t => {
  const launcher = await harness(t, route => route.fulfill(redirect("https://login.microsoftonline.com/")));
  const outlook = new Outlook(launcher, { launchVisible: launcher });
  const pending = outlook.signIn();
  await new Promise(resolve => setTimeout(resolve, 500));
  await launcher.contexts[0].close();
  assert.equal(await pending, "cancelled");
});

test("reports a network failure instead of a sign-in problem", async t => {
  const launcher = await harness(t, route => route.abort("internetdisconnected"));
  const outlook = new Outlook(launcher, { readyMs: 500 });
  await assert.rejects(outlook.check(), { code: "NETWORK_ERROR" });
  assert.equal(outlook.status().auth, "unknown");
  await outlook.close();
});

test("logout deletes only a browser profile", async t => {
  const dir = await mkdtemp(join(tmpdir(), "outlook-logout-"));
  t.after(() => rm(dir, { recursive: true, force: true }));
  const previous = process.env.OUTLOOK_PROFILE;
  t.after(() => { if (previous === undefined) delete process.env.OUTLOOK_PROFILE; else process.env.OUTLOOK_PROFILE = previous; });
  const profile = join(dir, "profile");
  await mkdir(join(profile, "Default"), { recursive: true });
  await writeFile(join(profile, "Local State"), "{}");
  process.env.OUTLOOK_PROFILE = profile;
  const outlook = new Outlook(async () => { throw new Error("not used"); });
  assert.deepEqual(await outlook.logout(), { profile_deleted: profile });
  assert.deepEqual(await readdir(dir), []);
  assert.equal(outlook.status().auth, "signed_out");
  await writeFile(join(dir, "notes.txt"), "keep me");
  process.env.OUTLOOK_PROFILE = dir;
  await assert.rejects(outlook.logout(), { code: "UNSAFE_PROFILE" });
  assert.deepEqual(await readdir(dir), ["notes.txt"]);
});
