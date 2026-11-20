import { test } from "node:test";
import assert from "node:assert/strict";
import { EventEmitter, getEventListeners } from "node:events";
import { Outlook } from "../dist/outlook.js";

function deferred() {
  let resolve;
  const promise = new Promise(done => { resolve = done; });
  return { promise, resolve };
}

function context() {
  const value = new EventEmitter();
  const page = new EventEmitter();
  page.isClosed = () => false;
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
