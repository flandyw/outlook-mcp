import { test } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { launchChromium } from "./chromium.mjs";
import { Calendar, parseCalendarDate } from "../dist/calendar.js";
import { Outlook } from "../dist/outlook.js";

const fixture = name => readFileSync(new URL("./fixtures/calendar/" + name, import.meta.url), "utf8");
const mailShell = fixture("mail-shell.html");
const dayView = fixture("day-view.html");

function calendarHtml({ mode = "url", drop = false } = {}) {
  return dayView.replaceAll("__MODE__", mode).replaceAll("__DROP__", drop ? "true" : "false");
}

test("parses English calendar headings and rejects impossible dates", () => {
  assert.equal(parseCalendarDate("Thursday, 12 November 2026"), "2026-11-12");
  assert.equal(parseCalendarDate("Wednesday, November 12, 2026"), "2026-11-12");
  assert.equal(parseCalendarDate("31 November 2026"), null);
  assert.equal(parseCalendarDate("no date here"), null);
});

test("rejects a bad range or unknown event without launching a browser", async () => {
  let launched = false;
  const outlook = new Outlook(async () => { launched = true; throw new Error("launched"); });
  const calendar = new Calendar(outlook);
  await assert.rejects(calendar.search({ start: "2026-11-31", limit: 5 }), { code: "INVALID_RANGE" });
  await assert.rejects(calendar.search({ start: "2026-02-29", limit: 5 }), { code: "INVALID_RANGE" });
  await assert.rejects(calendar.search({ start: "2026-11-20", end: "2026-11-01", limit: 5 }), { code: "INVALID_RANGE" });
  await assert.rejects(calendar.search({ end: "2026-11-12", limit: 5 }), { code: "INVALID_RANGE" });
  await assert.rejects(calendar.search({ start: "2026-11-12", end: "2026-11-26", limit: 5 }), { code: "INVALID_RANGE" });
  await assert.rejects(calendar.search({ start: "2026-11-12", query: "   ", limit: 5 }), { code: "INVALID_RANGE" });
  await assert.rejects(calendar.search({ start: "2026-11-12", limit: 0 }), { code: "INVALID_RANGE" });
  await assert.rejects(calendar.read("69d17491-5c2e-435d-b756-6c59752e2862"), { code: "UNKNOWN_EVENT" });
  assert.equal(launched, false);
});

async function start(t, respond) {
  const browser = await launchChromium();
  t.after(() => browser.close());
  const context = await browser.newContext();
  const leaks = [];
  context.on("request", request => {
    if (request.url().includes("example.com")) leaks.push(request.url());
  });
  await context.route(/^https:\/\/(outlook\.office\.com|login\.microsoftonline\.com)\//, route => respond(route));
  const outlook = new Outlook(async () => context, { readyMs: 3000 });
  t.after(() => outlook.close());
  return { context, outlook, calendar: new Calendar(outlook), leaks };
}

function serve(bodyForCalendar) {
  return route => {
    const url = new URL(route.request().url());
    if (url.hostname !== "outlook.office.com") {
      return route.fulfill({ contentType: "text/html", body: "<p>Sign in</p>" });
    }
    if (url.pathname.startsWith("/searchservice/")) {
      return route.fulfill({ contentType: "application/json", body: "{}" });
    }
    if (!route.request().isNavigationRequest()) return route.abort();
    if (url.pathname.startsWith("/calendar")) {
      const body = typeof bodyForCalendar === "function" ? bodyForCalendar(url) : bodyForCalendar;
      return route.fulfill({ contentType: "text/html", body });
    }
    return route.fulfill({ contentType: "text/html", body: mailShell });
  };
}

function calendarTab(context) {
  const pages = context.pages().filter(page => !page.isClosed() && page.url().includes("/calendar"));
  assert.ok(pages.length, "calendar tab was not opened");
  return pages[0];
}

test("lists and reads day-view events without touching mail or event actions", async t => {
  const { context, outlook, calendar, leaks } = await start(t, serve(calendarHtml()));
  const mail = await outlook.search("courier", 5);
  assert.equal(mail.messages.length, 1);
  const messageId = mail.messages[0].message_id;

  const today = await calendar.search({ limit: 10 });
  assert.equal(today.scope, "rendered_day_view");
  assert.deepEqual([today.start, today.end, today.query], ["2026-11-12", "2026-11-12", null]);
  assert.deepEqual(today.events.map(event => event.title), ["Specialist Maths", "Chemistry prac"]);
  assert.equal(today.events[0].all_day, false);
  assert.equal(today.events[0].location, "Room 4");
  assert.equal(today.events[0].start, "2026-11-12T09:00:00+11:00");
  assert.equal(today.events[0].end, "2026-11-12T10:15:00+11:00");
  assert.equal(today.events[0].date, "2026-11-12");
  assert.match(today.events[0].event_id, /^[0-9a-f-]{36}$/);
  assert.equal(new Set(today.events.map(event => event.event_id)).size, 2);
  assert.deepEqual(await calendarTab(context).evaluate(() => window.__clicks), {});
  assert.deepEqual(await calendarTab(context).evaluate(() => window.__actions), []);

  const room = await calendar.search({ start: "2026-11-12", query: "ROOM", limit: 5 });
  assert.deepEqual(room.events.map(event => event.title), ["Specialist Maths"]);
  const none = await calendar.search({ start: "2026-11-12", query: "orchestra", limit: 5 });
  assert.deepEqual(none.events, []);
  const one = await calendar.search({ start: "2026-11-12", limit: 1 });
  assert.deepEqual(one.events.map(event => event.title), ["Specialist Maths"]);

  const range = await calendar.search({ start: "2026-11-12", end: "2026-11-13", query: "hall", limit: 10 });
  assert.deepEqual(range.events.map(event => ({ title: event.title, all_day: event.all_day, location: event.location })), [
    { title: "Assembly", all_day: true, location: "Hall" },
  ]);
  assert.equal(range.events[0].start, "2026-11-13");
  assert.equal(range.events[0].end, "2026-11-13");

  const week = await calendar.search({ start: "2026-11-12", end: "2026-11-13", limit: 10 });
  assert.deepEqual(week.events.map(event => event.title), ["Specialist Maths", "Chemistry prac", "Assembly"]);
  const empty = await calendar.search({ start: "2026-11-14", limit: 5 });
  assert.deepEqual(empty.events, []);

  const mathsId = today.events[0].event_id;
  const maths = await calendar.read(mathsId);
  assert.equal(maths.title, "Specialist Maths");
  assert.equal(maths.location, "Room 4");
  assert.equal(maths.when, "Thursday, 12 November 2026 9:00 AM – 10:15 AM");
  assert.equal(maths.recurrence, "Occurs every Thursday until 20 November 2026");
  assert.equal(maths.all_day, false);
  assert.deepEqual(maths.attendees, [
    { name: "Alex Chen", response: "Accepted" },
    { name: "Sam Lee", response: "Tentative" },
    { name: "Jordan Lee", response: "Organizer" },
  ]);
  assert.equal(maths.truncated, false);
  assert.equal(maths.body_markdown, [
    "Bring the formula sheet.",
    "- Calculator",
    "Bring | Calculator",
    "\\# Not a heading",
    "(image: Seating plan)",
  ].join("\n\n"));
  assert.equal(maths.body_markdown.includes("example.com"), false);
  assert.equal(maths.body_markdown.includes("Hidden"), false);
  assert.deepEqual(await calendarTab(context).evaluate(() => window.__actions), []);
  assert.deepEqual(await calendarTab(context).evaluate(() => window.__clicks), { "evt-maths": 1 });

  const chem = await calendar.read(today.events[1].event_id);
  assert.equal(chem.body_markdown, "Safety glasses.");
  assert.equal(chem.recurrence, null);
  assert.deepEqual(chem.attendees, []);
  assert.equal(chem.location, "Lab 2");

  const assembly = await calendar.read(week.events[2].event_id);
  assert.equal(assembly.all_day, true);
  assert.equal(assembly.truncated, true);
  assert.equal(assembly.body_markdown, "a".repeat(12000));
  assert.equal(assembly.location, "Hall");

  // The id from the first search still resolves, and the mail tab was not replaced by the calendar.
  assert.equal((await calendar.read(mathsId)).title, "Specialist Maths");
  const message = await outlook.read(messageId);
  assert.equal(message.subject, "Delivery window");
  assert.equal(message.body_markdown, "The courier confirmed Friday.");
  assert.equal(message.received_at, "2026-11-11T10:42:00+11:00");

  const mailPage = context.pages().find(page => page.url().includes("/mail"));
  await mailPage.close();
  const again = await outlook.search("courier", 5);
  assert.equal(again.messages[0].subject, "Delivery window");
  assert.ok(context.pages().some(page => !page.isClosed() && page.url().includes("/mail")));
  assert.ok(context.pages().some(page => !page.isClosed() && page.url().includes("/calendar")));

  const span = await calendar.search({ start: "2026-11-12", end: "2026-11-25", limit: 10 });
  assert.deepEqual(span.events.map(event => event.title), ["Specialist Maths", "Chemistry prac", "Assembly"]);
  assert.deepEqual(leaks, []);
});

test("reaches a day with the next and previous buttons when the URL does not", async t => {
  const { calendar } = await start(t, serve(calendarHtml({ mode: "sticky" })));
  const day = await calendar.search({ start: "2026-11-13", limit: 5 });
  assert.deepEqual(day.events.map(event => event.title), ["Assembly"]);
  const back = await calendar.search({ start: "2026-11-12", limit: 5 });
  assert.deepEqual(back.events.map(event => event.title), ["Specialist Maths", "Chemistry prac"]);
});

test("reaches a day through the calendar date input", async t => {
  const { calendar } = await start(t, serve(calendarHtml({ mode: "input" })));
  const day = await calendar.search({ start: "2026-11-13", limit: 5 });
  assert.equal(day.events[0].title, "Assembly");
  assert.equal(day.events[0].all_day, true);
});

test("reads an English heading and aria-label when the structured fields are absent", async t => {
  const { calendar, leaks } = await start(t, serve(fixture("text-date.html")));
  const day = await calendar.search({ start: "2026-11-12", limit: 5 });
  assert.deepEqual(day.events.map(event => ({ title: event.title, start: event.start, end: event.end, all_day: event.all_day, location: event.location })), [
    { title: "Staff meeting", start: "2026-11-12", end: "2026-11-12", all_day: true, location: "Staff room" },
    { title: "Pastoral", start: "2026-11-12 8:30 AM", end: "2026-11-12 8:45 AM", all_day: false, location: "Courtyard" },
  ]);
  const pastoral = await calendar.read(day.events[1].event_id);
  assert.equal(pastoral.body_markdown, "Roll call");
  assert.equal(pastoral.when, "Thursday, 12 November 2026 8:30 AM – 8:45 AM");
  assert.equal(pastoral.recurrence, null);
  assert.deepEqual(pastoral.attendees, []);
  assert.deepEqual(leaks, []);
});

test("returns UI_CHANGED when the calendar layout cannot be identified", async t => {
  const { calendar } = await start(t, serve(fixture("unrecognized.html")));
  await assert.rejects(calendar.search({ start: "2026-11-12", limit: 5 }), { code: "UI_CHANGED" });
});

test("discards an event editor instead of saving it", async t => {
  const { context, calendar } = await start(t, serve(fixture("editor.html")));
  const day = await calendar.search({ start: "2026-11-12", limit: 5 });
  assert.equal(day.events[0].title, "Lab review");
  await assert.rejects(calendar.read(day.events[0].event_id), { code: "UI_CHANGED", message: /editor/ });
  assert.deepEqual(await calendarTab(context).evaluate(() => window.__actions), ["discard"]);
});

test("returns EVENT_CHANGED when the opened day no longer shows the event", async t => {
  let navigations = 0;
  const { calendar } = await start(t, serve(() => {
    navigations += 1;
    return calendarHtml({ drop: navigations > 1 });
  }));
  const day = await calendar.search({ start: "2026-11-12", limit: 5 });
  await assert.rejects(calendar.read(day.events[0].event_id), { code: "EVENT_CHANGED" });
});

test("returns AUTH_REQUIRED when the calendar redirects to Microsoft sign-in", async t => {
  const { calendar } = await start(t, route => {
    const url = new URL(route.request().url());
    if (url.hostname !== "outlook.office.com") {
      return route.fulfill({ contentType: "text/html", body: "<p>Sign in</p><input type='password' name='passwd'>" });
    }
    if (url.pathname.startsWith("/searchservice/")) return route.fulfill({ contentType: "application/json", body: "{}" });
    if (!route.request().isNavigationRequest()) return route.abort();
    if (url.pathname.startsWith("/calendar")) {
      return route.fulfill({ contentType: "text/html", body: "<script>location.replace(\"https://login.microsoftonline.com/common/oauth2/authorize\")</script>" });
    }
    return route.fulfill({ contentType: "text/html", body: mailShell });
  });
  await assert.rejects(calendar.search({ start: "2026-11-12", limit: 5 }), { code: "AUTH_REQUIRED" });
});

test("returns NETWORK_ERROR when the calendar cannot be reached", async t => {
  const { calendar } = await start(t, route => {
    const url = new URL(route.request().url());
    if (url.pathname.startsWith("/calendar")) return route.abort("internetdisconnected");
    if (url.pathname.startsWith("/searchservice/")) return route.fulfill({ contentType: "application/json", body: "{}" });
    if (!route.request().isNavigationRequest()) return route.abort();
    return route.fulfill({ contentType: "text/html", body: mailShell });
  });
  await assert.rejects(calendar.search({ start: "2026-11-12", limit: 5 }), { code: "NETWORK_ERROR" });
});

test("handles the current CalendarModule layout, restored week view and unlabelled event peek", async t => {
  const html = `<!doctype html><button aria-label="Day" onclick="document.querySelector('main').dataset.appSection='Surface_Day'">Day</button>
    <div data-app-section="CalendarModule">
      <div data-app-section="CalendarSurfaceNavigationToolbar"><button aria-label="Go to next day" onclick="document.querySelector('[data-column-date]').dataset.columnDate='2026-11-13'">Next</button></div>
      <main data-app-section="Surface_Week" aria-label="calendar view, current time: Thu 14:49">
        <div data-column-date="2026-11-12"></div><div data-app-section="calendar-view-0">
          <div role="button" aria-label="Lesson, 08:40 to 09:55, Thursday, November 12, 2026, S007, Busy, Categorized as Timetable" onclick="document.querySelector('#peek').hidden=false">Lesson</div>
          <div role="button" aria-label="Day 9, all day event, Thursday, November 12, 2026, Free">Day 9</div>
        </div>
      </main>
    </div>
    <div aria-live="polite">Loaded 2 events in 1 selected calendar.</div><div id="peek" data-app-section="CalendarItemPeek" hidden><span aria-label="Title">Lesson</span><span aria-label="Time">Thu 2026-11-12 08:40 - 09:55</span><div aria-label="Location">S007</div><button onclick="throw Error('must not edit')">Edit</button><button onclick="throw Error('must not delete')">Delete</button></div>`;
  const { calendar } = await start(t, serve(html));
  const result = await calendar.search({ start: "2026-11-13", limit: 10 });
  assert.equal(result.events.length, 2);
  const lesson = result.events.find(e => e.title === "Lesson");
  assert.equal(lesson.date, "2026-11-13");
  assert.equal(lesson.location, "S007");
  assert.equal(lesson.start, "2026-11-13 08:40");
  const details = await calendar.read(lesson.event_id);
  assert.equal(details.title, "Lesson");
  assert.equal(details.when, "Thu 2026-11-12 08:40 - 09:55");
  assert.equal(details.location, "S007");
});
