import { test } from "node:test";
import assert from "node:assert/strict";
import { EventEmitter, getEventListeners } from "node:events";
import { Outlook, MailError, errorInfo } from "../dist/outlook.js";

function deferred() {
  let resolve;
  const promise = new Promise(done => { resolve = done; });
  return { promise, resolve };
}

function context() {
  const value = new EventEmitter();
  const page = new EventEmitter();
  page.isClosed = () => false;
  page.url = () => "https://outlook.office.com/mail/";
  page.setDefaultTimeout = () => {};
  value.pages = () => [page];
  value.closed = false;
  value.close = async () => { value.closed = true; value.emit("close"); };
  return value;
}

test("queue deadlines include waiting, drop expired/cancelled work and remove abort listeners", async t => {
  const options = { operationMs: 2000 };
  const outlook = new Outlook(async () => context(), options);
  outlook.ensureMailbox = async () => {};
  t.after(() => outlook.close());
  const started = deferred();
  const release = deferred();
  const first = outlook.useBrowser(async () => { started.resolve(); await release.promise; });
  await started.promise;
  options.operationMs = 40;
  let ran = 0;
  const expired = outlook.useBrowser(async () => { ran++; });
  await assert.rejects(expired, { code: "TIMEOUT" });
  assert.equal(ran, 0);
  assert.equal(outlook.status().browser, "running", "queued expiry must not reset the active operation's browser");
  const cancelled = new AbortController();
  const waiting = outlook.useBrowser(async () => { ran++; }, cancelled.signal);
  cancelled.abort();
  await assert.rejects(waiting, { code: "CANCELLED" });
  release.resolve();
  await first;
  options.operationMs = 2000;
  await outlook.useBrowser(async () => {});
  assert.equal(ran, 0);
  assert.equal(getEventListeners(cancelled.signal, "abort").length, 0);
  const already = new AbortController();
  already.abort();
  await assert.rejects(outlook.useBrowser(async () => { ran++; }, already.signal), { code: "CANCELLED" });
  assert.equal(ran, 0);
});

test("transient operations get one clean-view retry; auth, validation and partial effects do not", async t => {
  for (const code of ["UI_CHANGED", "TIMEOUT", "NETWORK_ERROR", "BROWSER_RESTARTED", "DOWNLOAD_FAILED", "AUTH_REQUIRED", "INVALID_RANGE", "MESSAGE_CHANGED", "STORAGE_FULL"]) {
    const outlook = new Outlook(async () => context(), { operationMs: 2000 });
    t.after(() => outlook.close());
    let clean = 0, calls = 0;
    outlook.ensureMailbox = async (page, force) => { if (force) clean++; };
    outlook.sessionLost = async () => false;
    const retryable = ["UI_CHANGED", "TIMEOUT", "NETWORK_ERROR", "BROWSER_RESTARTED", "DOWNLOAD_FAILED"].includes(code);
    const work = () => outlook.useBrowser(async () => {
      if (++calls === 1) throw new MailError(code, "fixture failure");
      return "recovered";
    });
    if (retryable) {
      assert.equal(await work(), "recovered", code);
      assert.equal(calls, 2, code);
      assert.equal(clean, 1, code);
    } else {
      await assert.rejects(work(), { code });
      assert.equal(calls, 1, code);
      assert.equal(clean, 0, code);
    }
  }
  assert.equal(errorInfo(new MailError("UI_CHANGED", "changed")).retryable, true);
  const outlook = new Outlook(async () => context());
  t.after(() => outlook.close());
  outlook.ensureMailbox = async () => {};
  outlook.sessionLost = async () => false;
  let calls = 0;
  await assert.rejects(outlook.useBrowser(async () => {
    calls++;
    outlook.preventRetry();
    throw new MailError("UI_CHANGED", "already saved");
  }), { code: "UI_CHANGED" });
  assert.equal(calls, 1);
});

test("a second transient failure is final and clean-view recovery closes calendar tabs", async t => {
  const outlook = new Outlook(async () => context());
  t.after(() => outlook.close());
  outlook.ensureMailbox = async () => {};
  outlook.sessionLost = async () => false;
  let calls = 0;
  const calendar = new EventEmitter();
  let closed = false;
  calendar.close = async () => { closed = true; calendar.emit("close"); };
  outlook.holdPage(calendar);
  await assert.rejects(outlook.useBrowser(async () => {
    if (++calls === 2) assert.equal(closed, true);
    throw new MailError("UI_CHANGED", "still changed");
  }), { code: "UI_CHANGED", message: "still changed" });
  assert.equal(calls, 2);
});

test("recovery stays within the original deadline and respects cancellation", async t => {
  for (const cancelled of [false, true]) {
    const outlook = new Outlook(async () => context(), { operationMs: cancelled ? 2000 : 50 });
    t.after(() => outlook.close());
    const cleaning = deferred(), release = deferred();
    outlook.ensureMailbox = async (page, force) => {
      if (force) { cleaning.resolve(); await release.promise; }
    };
    outlook.sessionLost = async () => false;
    const controller = new AbortController();
    let calls = 0;
    const operation = outlook.useBrowser(async () => {
      calls++;
      throw new MailError("UI_CHANGED", "retry me");
    }, controller.signal);
    const rejected = assert.rejects(operation, { code: cancelled ? "CANCELLED" : "TIMEOUT" });
    await cleaning.promise;
    if (cancelled) controller.abort();
    await rejected;
    release.resolve();
    await outlook.useBrowser(async () => {});
    assert.equal(calls, 1, "no second work after cancellation/deadline during recovery");
  }
});

test("launch is inside the enqueue deadline and a late browser is closed rather than adopted", async t => {
  const launched = deferred();
  const late = context();
  const closed = deferred();
  late.once("close", closed.resolve);
  let calls = 0;
  let ran = false;
  const outlook = new Outlook(async () => ++calls === 1 ? launched.promise : context(), { operationMs: 40 });
  outlook.ensureMailbox = async () => {};
  t.after(() => outlook.close());
  await assert.rejects(outlook.useBrowser(async () => { ran = true; }), { code: "TIMEOUT" });
  assert.equal(outlook.status().browser, "stopped");
  launched.resolve(late);
  await closed.promise;
  assert.equal(ran, false);
  await outlook.useBrowser(async () => {});
  assert.equal(calls, 2);
});

test("active cancellation stops at checkpoints without releasing the queue before cleanup", async t => {
  const outlook = new Outlook(async () => context(), { operationMs: 2000 });
  outlook.ensureMailbox = async () => {};
  t.after(() => outlook.close());
  const started = deferred();
  const release = deferred();
  const controller = new AbortController();
  const actions = [];
  const active = outlook.useBrowser(async () => {
    try {
      started.resolve();
      await release.promise;
      outlook.checkpoint();
      actions.push("must not run");
    } finally { actions.push("cleanup"); }
  }, controller.signal);
  await started.promise;
  controller.abort();
  await assert.rejects(active, { code: "CANCELLED" });
  const following = outlook.useBrowser(async () => { actions.push("next"); });
  assert.deepEqual(actions, []);
  release.resolve();
  await following;
  assert.deepEqual(actions, ["cleanup", "next"]);
  assert.equal(outlook.status().browser, "running");
  assert.equal(getEventListeners(controller.signal, "abort").length, 0);
});
