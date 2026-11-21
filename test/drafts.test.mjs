import { test } from "node:test";
import assert from "node:assert/strict";
import { launchChromium } from "./chromium.mjs";
import { composeFixture } from "./fixtures/mail-compose.mjs";
import { Outlook, MailError, draftInput } from "../dist/outlook.js";
import { toolContracts } from "../dist/tool-contracts.js";

process.env.OUTLOOK_DRAFTS = "1";
const input = { to: ["new@example.com"], cc: ["cc@example.com"], subject: "A draft", body: "Line one\n<img src=x onerror=alert(1)>\nLine three" };
async function fixture(t, options = {}) {
  const browser = await launchChromium();
  t.after(() => browser.close());
  const context = await browser.newContext();
  await context.route("https://outlook.office.com/**", route => route.fulfill({ contentType: "text/html", body: composeFixture(options) }));
  const outlook = new Outlook(async () => context);
  const listed = await outlook.list();
  const page = context.pages()[0];
  return { outlook, page, id: listed.messages[0].message_id, state: () => page.evaluate(() => window.fixture) };
}

test("new drafts are saved and closed with resolved recipients and literal multiline text", async t => {
  const { outlook, page, state } = await fixture(t);
  const result = await outlook.createDraft(input);
  toolContracts.create_draft.outputSchema.parse(result);
  assert.deepEqual(result.to, input.to);
  assert.deepEqual(result.cc, input.cc);
  assert.equal(result.reply_to_message_id, null);
  assert.equal(result.draft, true);
  assert.ok(Number.isFinite(Date.parse(result.saved_at)));
  const observed = await state();
  assert.equal(observed.sends, 0);
  assert.equal(observed.closed, 1);
  assert.equal(observed.discards, 0);
  assert.equal(observed.saved[0].body, input.body);
  assert.match(observed.saved[0].html, /&lt;img/);
  assert.equal(await page.locator('[data-testid="compose-editor"]').count(), 0);
  for (const body of ["Trailing\n", "\n\n", " \tindented\n  two "]) {
    assert.equal((await outlook.createDraft({ ...input, body })).draft, true);
  }
});

test("reply and reply-all preserve prefilled recipients and support empty additions/body", async t => {
  for (const reply_mode of ["reply", "reply_all"]) {
    const { outlook, state, id } = await fixture(t, { unread: true });
    const result = await outlook.createDraft({ to: ["extra@example.com"], cc: ["added@example.com"], body: "Reply\ntext", reply_to_message_id: id, reply_mode });
    assert.equal(result.subject, "Re: Original");
    assert.equal(result.reply_to_message_id, id);
    assert.deepEqual(result.to, reply_mode === "reply" ? ["extra@example.com", "ada@example.com"] : ["extra@example.com", "team@example.com", "ada@example.com"]);
    assert.deepEqual(result.cc, reply_mode === "reply" ? ["added@example.com"] : ["added@example.com", "copy@example.com"]);
    assert.equal((await state()).sends, 0);
    assert.deepEqual((await state()).modes, [reply_mode]);
    assert.equal((await state()).restored, 1);
    const empty = await outlook.createDraft({ to: [], body: "", reply_to_message_id: id });
    assert.deepEqual(empty.to, ["ada@example.com"]);
    assert.equal((await state()).saved[1].body, "");
    assert.equal((await state()).restored, 2);
  }
});

test("autosave must produce a fresh saved indication", async t => {
  const { outlook, state } = await fixture(t, { autoSave: true, staleSaved: true });
  await outlook.createDraft(input);
  assert.equal((await state()).saved.length, 1);
  assert.equal((await state()).saved[0].body, input.body);
  assert.equal((await state()).sends, 0);
});

test("reply and reply-all wait for asynchronously rendered controls before the guarded click", async t => {
  for (const reply_mode of ["reply", "reply_all"]) {
    const { outlook, state, id } = await fixture(t, { replyDelay: 500, unread: true });
    const result = await outlook.createDraft({ to: [], body: "Delayed reply", reply_to_message_id: id, reply_mode });
    assert.equal(result.draft, true);
    const observed = await state();
    assert.deepEqual(observed.modes, [reply_mode]);
    assert.equal(observed.opens, 1);
    assert.equal(observed.saved.length, 1);
    assert.equal(observed.restored, 1);
    assert.equal(observed.sends, 0);
  }
});

test("delayed ambiguous reply controls are rejected without opening a composer or clicking Send", async t => {
  const { outlook, state, id } = await fixture(t, { replyDelay: 500, ambiguousReply: true });
  await assert.rejects(outlook.createDraft({ to: [], body: "", reply_to_message_id: id }), { code: "DRAFT_UNCERTAIN" });
  assert.equal((await state()).opens, 0);
  assert.equal((await state()).sends, 0);
});

test("waiting for reply controls respects cancellation and never opens a composer", async t => {
  const { outlook, page, state, id } = await fixture(t, { replyDelay: 10000, unread: true });
  const controller = new AbortController();
  const pending = outlook.createDraft({ to: [], body: "", reply_to_message_id: id }, controller.signal);
  const rejected = assert.rejects(pending, { code: "CANCELLED" });
  await page.getByRole("heading", { name: "Original", exact: true }).waitFor({ state: "visible" });
  controller.abort();
  await rejected;
  await outlook.useBrowser(async () => {});
  assert.equal((await state()).opens, 0);
  assert.equal((await state()).restored, 1);
  assert.equal((await state()).sends, 0);
});

test("unresolved recipients fail clearly and discard without opening a second composer", async t => {
  const { outlook, state } = await fixture(t, { unresolved: true });
  await assert.rejects(outlook.createDraft(input), { code: "DRAFT_RECIPIENT_UNRESOLVED" });
  const observed = await state();
  assert.equal(observed.opens, 1);
  assert.equal(observed.discards, 1);
  assert.equal(observed.saved.length, 0);
  assert.equal(observed.sends, 0);
});

test("send guard blocks misleading Save labels and ambiguous controls", async t => {
  for (const options of [{ poisonSave: true }, { ambiguousSave: true }]) {
    const { outlook, state } = await fixture(t, options);
    await assert.rejects(outlook.createDraft(input), { code: options.poisonSave ? "DRAFT_SEND_BLOCKED" : "DRAFT_UNCERTAIN" });
    assert.equal((await state()).opens, 1);
    assert.equal((await state()).sends, 0);
    assert.equal((await state()).discards, 1);
  }
});

test("stale saved state and failure to close return uncertainty without retry or deleting a saved draft", async t => {
  for (const options of [{ noSave: true, staleSaved: true }, { noClose: true }, { replaceBody: true }]) {
    const { outlook, state } = await fixture(t, options);
    await assert.rejects(outlook.createDraft(input), { code: "DRAFT_UNCERTAIN" });
    const observed = await state();
    assert.equal(observed.opens, 1);
    assert.equal(observed.sends, 0);
    assert.equal(observed.discards, options.noClose ? 0 : 1);
    assert.equal(observed.saved.length, options.noSave ? 0 : 1);
  }
});

test("transient failures after opening are never replayed", async t => {
  const { outlook, state } = await fixture(t);
  const click = outlook.draftClick.bind(outlook);
  outlook.draftClick = async target => {
    if (await target.getAttribute("id") === "save") throw new MailError("UI_CHANGED", "fixture transient failure");
    return click(target);
  };
  await assert.rejects(outlook.createDraft(input), { code: "UI_CHANGED" });
  assert.equal((await state()).opens, 1);
  assert.equal((await state()).discards, 1);
  assert.equal((await state()).sends, 0);
});

test("cancellation after opening discards the unfinished composer before releasing the queue", async t => {
  const { outlook, page, state } = await fixture(t, { unresolved: true, confirmDiscard: true });
  const controller = new AbortController();
  const pending = outlook.createDraft(input, controller.signal);
  const rejected = assert.rejects(pending, { code: "CANCELLED" });
  await page.waitForFunction(() => window.fixture.opens === 1);
  controller.abort();
  await rejected;
  await outlook.useBrowser(async () => {});
  assert.equal((await state()).discards, 1);
  assert.equal((await state()).sends, 0);
  assert.equal((await state()).saved.length, 0);
});

test("draft validation and opt-in fail before opening any composer", async t => {
  const { outlook, state } = await fixture(t);
  delete process.env.OUTLOOK_DRAFTS;
  assert.throws(() => outlook.createDraft(input), { code: "DRAFTS_DISABLED" });
  process.env.OUTLOOK_DRAFTS = "1";
  for (const bad of [{ ...input, to: [] }, { ...input, to: ["not an email"] }, { ...input, subject: undefined },
    { ...input, to: Array(51).fill("a@example.com") }, { ...input, cc: Array(51).fill("a@example.com") },
    { ...input, subject: "x".repeat(256) }, { ...input, body: "x".repeat(20001) }]) {
    assert.throws(() => outlook.createDraft(bad), { code: "INVALID_DRAFT" });
  }
  assert.equal(draftInput.parse(input).reply_mode, "reply");
  assert.equal((await state()).opens, 0);
});

test("an existing composer is left untouched", async t => {
  const { outlook, page, state } = await fixture(t);
  await page.getByRole("button", { name: "New mail", exact: true }).click();
  await assert.rejects(outlook.createDraft(input), { code: "DRAFT_BUSY" });
  assert.equal((await state()).opens, 1);
  assert.equal((await state()).discards, 0);
  assert.equal((await state()).sends, 0);
});

test("lost supplied or prefilled recipient pills fail before saving", async t => {
  for (const reply of [false, true]) {
    const { outlook, id, state } = await fixture(t, { replaceRecipients: true });
    await assert.rejects(outlook.createDraft({ ...input, to: ["first@example.com", "second@example.com"],
      ...(reply ? { reply_to_message_id: id } : {}) }), { code: "DRAFT_RECIPIENT_UNRESOLVED" });
    assert.equal((await state()).opens, 1);
    assert.equal((await state()).saved.length, 0);
    assert.equal((await state()).sends, 0);
  }
});

test("replies permit empty To additions and a CC-only prefilled draft", async t => {
  const { outlook, id } = await fixture(t, { replyCcOnly: true });
  const result = await outlook.createDraft({ to: [], body: "", reply_to_message_id: id });
  assert.deepEqual(result.to, []);
  assert.deepEqual(result.cc, ["copy@example.com"]);
});

test("cancellation after save closes but never discards the verified draft", async t => {
  const { outlook, state } = await fixture(t);
  const controller = new AbortController();
  const click = outlook.draftClick.bind(outlook);
  let cancelled = false;
  outlook.draftClick = async target => {
    if (!cancelled && await target.getAttribute("id") === "close") {
      cancelled = true;
      controller.abort();
      throw new MailError("CANCELLED", "fixture cancellation");
    }
    return click(target);
  };
  await assert.rejects(outlook.createDraft(input, controller.signal), { code: "CANCELLED" });
  await outlook.useBrowser(async () => {});
  const observed = await state();
  assert.equal(observed.opens, 1);
  assert.equal(observed.saved.length, 1);
  assert.equal(observed.closed, 1);
  assert.equal(observed.discards, 0);
  assert.equal(observed.sends, 0);
});
