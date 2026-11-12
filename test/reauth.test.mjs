import { test } from "node:test";
import { strict as assert } from "node:assert";
import { execFileSync, spawnSync } from "node:child_process";
import { mkdtemp, mkdir, readdir, rm, symlink, writeFile, chmod } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { launchChromium } from "./chromium.mjs";
import { inspectPasswordFile, reauthenticate } from "../dist/reauth.js";

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

const reportKeys = ["configured", "hints", "message", "path", "problem", "valid"];
const canary = "canary-password-value-do-not-print";

async function fixtureDir(t) {
  const dir = await mkdtemp(join(tmpdir(), "outlook-password-"));
  t.after(async () => {
    const names = await readdir(dir).catch(() => []);
    await Promise.all(names.map(name => chmod(join(dir, name), 0o700).catch(() => {})));
    await rm(dir, { recursive: true, force: true });
  });
  return dir;
}

async function privateFile(path, contents, mode = 0o600) {
  await writeFile(path, contents, { mode });
  await chmod(path, mode);
}

test("checks the configured password file without reading secrets into the report", async t => {
  const dir = await fixtureDir(t);
  assert.deepEqual(inspectPasswordFile(undefined), {
    configured: false, valid: false, path: null, problem: null, hints: [], message: null,
  });
  assert.equal(inspectPasswordFile("  ").configured, false);

  const missing = join(dir, "alpha-token-value");
  await privateFile(join(dir, "alpha-token"), canary + "\n");
  await privateFile(join(dir, "alpha-token-value-bak"), canary + "\n");
  await privateFile(join(dir, "notes.txt"), canary + "\n");
  await mkdir(join(dir, "alpha-token-value-old"));
  const missingReport = inspectPasswordFile(missing);
  assert.equal(missingReport.configured, true);
  assert.equal(missingReport.valid, false);
  assert.equal(missingReport.problem, "missing");
  assert.equal(missingReport.path, missing);
  assert.ok(missingReport.hints.includes(join(dir, "alpha-token")));
  assert.ok(missingReport.hints.includes(join(dir, "alpha-token-value-bak")));
  assert.equal(missingReport.hints.includes(join(dir, "notes.txt")), false);
  assert.equal(missingReport.hints.includes(join(dir, "alpha-token-value-old")), false);
  assert.match(missingReport.message, /Password file misconfigured/);
  assert.match(missingReport.message, /is missing/);
  assert.match(missingReport.message, /not used automatically/);
  assert.equal(missingReport.message.includes(join(dir, "alpha-token-value-bak")), true);
  assert.equal(missingReport.message.includes(join(dir, "notes.txt")), false);
  assert.equal(missingReport.message.includes(canary), false);
  assert.deepEqual(Object.keys(missingReport).sort(), reportKeys);
  const absentDir = inspectPasswordFile(join(dir, "no-such-directory", "secret"));
  assert.equal(absentDir.problem, "missing");
  assert.deepEqual(absentDir.hints, []);

  const ready = join(dir, "ready");
  await privateFile(ready, canary + "\n", 0o600);
  const readyReport = inspectPasswordFile(ready);
  assert.equal(readyReport.valid, true);
  assert.equal(readyReport.problem, null);
  assert.equal(readyReport.message, null);
  assert.equal(JSON.stringify(readyReport).includes(canary), false);

  const ownerOnly = join(dir, "owner-only");
  await privateFile(ownerOnly, "x\n", 0o400);
  assert.equal(inspectPasswordFile(ownerOnly).valid, true);

  const loose = join(dir, "loose");
  await privateFile(loose, canary + "\n", 0o644);
  const looseReport = inspectPasswordFile(loose);
  assert.equal(looseReport.problem, "loose_permissions");
  assert.match(looseReport.message, /loose permissions/);
  assert.match(looseReport.message, /0644/);
  assert.equal(looseReport.message.includes(canary), false);
  await chmod(loose, 0o620);
  assert.equal(inspectPasswordFile(loose).problem, "loose_permissions");

  const blank = join(dir, "blank");
  await privateFile(blank, "\n");
  assert.equal(inspectPasswordFile(blank).problem, "empty");
  await privateFile(blank, "");
  assert.match(inspectPasswordFile(blank).message, /is empty/);

  const unread = join(dir, "unread");
  await privateFile(unread, canary + "\n", 0o000);
  const unreadReport = inspectPasswordFile(unread);
  assert.equal(unreadReport.problem, "unreadable");
  assert.match(unreadReport.message, /unreadable/);
  assert.equal(unreadReport.message.includes(canary), false);

  const huge = join(dir, "huge");
  await privateFile(huge, "a".repeat(4097));
  const hugeReport = inspectPasswordFile(huge);
  assert.equal(hugeReport.problem, "too_large");
  assert.equal(hugeReport.message.includes("aaa"), false);

  const folder = join(dir, "folder");
  await mkdir(folder);
  assert.match(inspectPasswordFile(folder).message, /not a regular file/);
  const link = join(dir, "link");
  await symlink(ready, link);
  assert.equal(inspectPasswordFile(link).problem, "not_a_file");
});

test("does not sign in with a similarly named password file", async t => {
  const dir = await fixtureDir(t);
  const configured = join(dir, "mailbox-credential");
  await privateFile(join(dir, "mailbox-credentials"), canary + "\n");
  const browser = await launchChromium();
  t.after(() => browser.close());
  const page = await browser.newPage();
  await page.route("https://login.microsoftonline.com/**", route => route.fulfill({
    contentType: "text/html",
    body: `<form><p>fixture@example.com</p><input type="password" name="passwd"><input type="submit" id="idSIButton9"></form>
      <script>document.forms[0].onsubmit = (event) => {
        event.preventDefault();
        location.href = 'https://outlook.office.com/mail/';
      };</script>`,
  }));
  await page.route("https://outlook.office.com/**", route => route.fulfill({ body: "Mailbox fixture" }));
  await page.goto("https://login.microsoftonline.com/");
  await assert.rejects(reauthenticate(page, "fixture@example.com", configured), error => {
    assert.match(error.message, /Password file misconfigured/);
    assert.match(error.message, /is missing/);
    assert.match(error.message, /mailbox-credentials/);
    assert.match(error.message, /not used automatically/);
    assert.equal(error.message.includes(canary), false);
    return true;
  });
  assert.equal(page.url(), "https://login.microsoftonline.com/");
});

test("rejects a loose, empty, or unreadable password file before typing it", async t => {
  const dir = await fixtureDir(t);
  const loose = join(dir, "loose");
  await privateFile(loose, canary + "\n", 0o644);
  const empty = join(dir, "empty");
  await privateFile(empty, "\n");
  const unread = join(dir, "unread");
  await privateFile(unread, canary, 0o000);
  for (const [file, pattern] of [[loose, /loose permissions/], [empty, /is empty/], [unread, /unreadable/]]) {
    await assert.rejects(reauthenticate({}, "user@example.com", file), error => {
      assert.match(error.message, pattern);
      assert.equal(error.message.includes(canary), false);
      return true;
    });
  }
});

test("does not block when the password path is a fifo", t => {
  const dir = execFileSync("mktemp", ["-d"], { encoding: "utf8" }).trim();
  const fifo = join(dir, "fifo");
  t.after(() => rm(dir, { recursive: true, force: true }));
  execFileSync("mkfifo", [fifo]);
  const result = spawnSync(process.execPath, ["--input-type=module", "-e", `
    import { inspectPasswordFile, reauthenticate } from "./dist/reauth.js";
    const report = inspectPasswordFile(process.env.PASSWORD_PATH);
    if (report.problem !== "not_a_file") process.exit(2);
    try {
      await reauthenticate({}, "user@example.com", process.env.PASSWORD_PATH);
      process.exit(3);
    } catch (error) {
      if (!/not a regular file/.test(error.message)) process.exit(4);
    }
  `], {
    timeout: 10000,
    encoding: "utf8",
    env: { ...process.env, PASSWORD_PATH: fifo, OUTLOOK_PASSWORD_FILE: "" },
  });
  assert.equal(result.error, undefined, result.stderr || result.error?.message);
  assert.equal(result.status, 0, result.stderr);
});
