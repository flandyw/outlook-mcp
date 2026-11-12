import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, mkdir, rm, writeFile, readdir, chmod } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { launchChromium } from "./chromium.mjs";
import { Outlook } from "../dist/outlook.js";

// Tests that change sign-in settings share one queue so they do not clobber each other's environment.
let envQueue = Promise.resolve();
function withEnv(updates, body) {
  const run = envQueue.then(async () => {
    const saved = new Map(Object.keys(updates).map(key => [key, process.env[key]]));
    for (const [key, value] of Object.entries(updates)) {
      if (value === undefined) delete process.env[key];
      else process.env[key] = value;
    }
    try {
      await body();
    } finally {
      for (const [key, value] of saved) {
        if (value === undefined) delete process.env[key];
        else process.env[key] = value;
      }
    }
  });
  envQueue = run.then(() => {}, () => {});
  return run;
}

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
  await withEnv({ OUTLOOK_PASSWORD_FILE: undefined, OUTLOOK_USERNAME: undefined }, async () => {
    const changes = [];
    const noted = [];
    const outlook = new Outlook(launcher, {
      onAuthChange: (state, detail) => {
        changes.push(state);
        if (detail?.message) noted.push(detail.message);
      },
      readyMs: 1000,
    });
    const expectExpired = error => {
      assert.equal(error.code, "AUTH_REQUIRED");
      assert.match(error.message, /Microsoft session expired/);
      assert.match(error.message, /npm run login/);
      assert.doesNotMatch(error.message, /Password file/);
      assert.doesNotMatch(error.message, /login window is already open/);
      return true;
    };
    await assert.rejects(outlook.search("x", 5), expectExpired);
    await assert.rejects(outlook.search("x", 5), expectExpired);
    assert.deepEqual(changes, ["signed_out"]);
    assert.match(noted[0], /Microsoft session expired/);
    assert.equal(outlook.status().auth, "signed_out");
    assert.equal(outlook.status().last_error.code, "AUTH_REQUIRED");
    assert.match(outlook.status().last_error.message, /Microsoft session expired/);
    await outlook.close();
  });
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
  await assert.rejects(outlook.search("x", 5), { code: "AUTH_REQUIRED", message: /login window is already open/ });
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

test("names a misconfigured password file instead of a vague sign-in failure", async t => {
  const dir = await mkdtemp(join(tmpdir(), "outlook-password-"));
  t.after(() => rm(dir, { recursive: true, force: true }));
  const configured = join(dir, "alpha-token-value");
  const hint = join(dir, "alpha-token-value-bak");
  const canary = "session-canary-do-not-print";
  await writeFile(hint, canary + "\n", { mode: 0o600 });
  await chmod(hint, 0o600);
  const launcher = await harness(t, route => route.request().url().startsWith("https://outlook.office.com/")
    ? route.fulfill(redirect("https://login.microsoftonline.com/common/oauth2/authorize"))
    : route.fulfill({ contentType: "text/html", body: '<p>fixture@example.com</p><input type="password" name="passwd"><input type="submit" id="idSIButton9">' }));
  await withEnv({ OUTLOOK_USERNAME: "fixture@example.com", OUTLOOK_PASSWORD_FILE: configured }, async () => {
    const noted = [];
    const outlook = new Outlook(launcher, {
      onAuthChange: (_state, detail) => { if (detail?.message) noted.push(detail.message); },
      readyMs: 1000,
    });
    const expectFile = error => {
      assert.equal(error.code, "AUTH_REQUIRED");
      assert.match(error.message, /Password file misconfigured/);
      assert.match(error.message, /is missing/);
      assert.equal(error.message.includes(hint), true);
      assert.match(error.message, /not used automatically/);
      assert.doesNotMatch(error.message, /Automatic sign-in failed/);
      assert.doesNotMatch(error.message, /cooling down/);
      assert.equal(error.message.includes(canary), false);
      return true;
    };
    await assert.rejects(outlook.search("x", 5), expectFile);
    await assert.rejects(outlook.search("x", 5), expectFile);
    assert.equal(noted.length, 1);
    assert.match(noted[0], /Password file misconfigured/);
    assert.equal(noted[0].includes(canary), false);
    await outlook.close();
  });
});

test("distinguishes a failed automatic sign-in from the cooldown that follows", async t => {
  const dir = await mkdtemp(join(tmpdir(), "outlook-password-"));
  t.after(() => rm(dir, { recursive: true, force: true }));
  const file = join(dir, "password");
  const password = "fixture-password";
  await writeFile(file, password + "\n", { mode: 0o600 });
  await chmod(file, 0o600);
  const launcher = await harness(t, route => {
    if (route.request().url().startsWith("https://outlook.office.com/")) {
      return route.fulfill(redirect("https://login.microsoftonline.com/common/oauth2/authorize"));
    }
    return route.fulfill({ contentType: "text/html", body: `<form><p>fixture@example.com</p>
      <input type="password" name="passwd"><input type="submit" id="idSIButton9"></form>
      <script>document.forms[0].onsubmit = (event) => { event.preventDefault(); location.href = "about:blank"; };</script>` });
  });
  await withEnv({ OUTLOOK_USERNAME: "fixture@example.com", OUTLOOK_PASSWORD_FILE: file }, async () => {
    const outlook = new Outlook(launcher, { readyMs: 1000 });
    await assert.rejects(outlook.search("x", 5), error => {
      assert.equal(error.code, "AUTH_REQUIRED");
      assert.match(error.message, /Automatic sign-in failed/);
      assert.match(error.message, /npm run login/);
      assert.doesNotMatch(error.message, /Password file misconfigured/);
      assert.equal(error.message.includes(password), false);
      return true;
    });
    await assert.rejects(outlook.search("x", 5), error => {
      assert.equal(error.code, "AUTH_REQUIRED");
      assert.match(error.message, /cooling down/);
      assert.match(error.message, /npm run login/);
      assert.doesNotMatch(error.message, /Password file misconfigured/);
      assert.equal(error.message.includes(password), false);
      return true;
    });
    await outlook.close();
  });
});

test("discards the browser after a Playwright navigation timeout", async t => {
  const launcher = await harness(t, route => route.fulfill({ contentType: "text/html", body: mailboxPage }));
  const outlook = new Outlook(launcher);
  await outlook.check();
  const page = launcher.contexts[0].pages()[0];
  // Trigger the same native TimeoutError as page.goto without a slow live network.
  const timeout = await page.locator("#missing").waitFor({ timeout: 20 }).catch(error => error);
  page.goto = async () => { throw timeout; };
  await assert.rejects(outlook.openSignIn(), { code: "TIMEOUT" });
  assert.equal(outlook.status().browser, "stopped");
  await outlook.check();
  assert.equal(launcher.count, 2);
  await outlook.close();
});
