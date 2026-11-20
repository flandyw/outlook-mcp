import { randomUUID } from "node:crypto";
import type { BrowserContext, Page } from "playwright";
import { mailbox, outlookOrigins, safeOrigin } from "./browser.js";
import { formatBody, type BodyPart } from "./format.js";
import { MailError, SessionLostError, errorInfo, waitForSurface, type Outlook } from "./outlook.js";
import { isMicrosoftLogin } from "./reauth.js";

const signInHelp = "Run npm run login on the server host to sign in again.";
const idTtlMs = 30 * 60_000;
// 14 calendar days, inclusive: the span between the first and last date is 13.
const maxSpanDays = 13;

const months = ["january", "february", "march", "april", "may", "june", "july", "august", "september", "october", "november", "december"];

// Selectors cover the observed Outlook CalendarModule / Surface_Day layout and fixtures:
// a calendar surface, a date heading, event buttons in the day grid, and a details peek.
// A missing surface or grid is UI_CHANGED rather than a guess.
const selectors = {
  surface: '[data-testid="calendar-surface"], [data-app-section="CalendarModule"]',
  date: '[data-testid="calendar-date"]',
  event: '[data-testid="calendar-event"], [role="grid"] [role="button"], [role="grid"] button, [data-app-section="Surface_Day"] [role="button"][aria-label]',
  title: "[data-testid='event-title']",
  start: "[data-testid='event-start']",
  end: "[data-testid='event-end']",
  location: "[data-testid='event-location']",
  peek: '[data-testid="event-peek"], [role="dialog"], [data-app-section="CalendarItemPeek"]',
  heading: '[data-testid="event-heading"], [aria-label="Title"], [role="heading"], h1, h2',
  when: "[data-testid='event-when'], [aria-label='Time']",
  where: "[data-testid='event-where'], [aria-label='Location']",
  recurrence: "[data-testid='event-recurrence']",
  attendee: "[data-testid='event-attendees'] [role='listitem']",
  notes: "[data-testid='event-notes'], [aria-label='Description'], [aria-label='Notes']",
};

type RawEvent = {
  index: number;
  sourceId: string | null;
  title: string;
  start: string;
  end: string;
  allDay: boolean;
  location: string | null;
};

type ListedEvent = RawEvent & { date: string; fingerprint: string };

type EventRef = {
  date: string;
  title: string;
  start: string;
  end: string;
  allDay: boolean;
  location: string | null;
  fingerprint: string;
  expires: number;
};

export type EventSummary = {
  event_id: string;
  title: string;
  start: string;
  end: string;
  all_day: boolean;
  location: string | null;
  date: string;
};

export type EventDetails = EventSummary & {
  when: string | null;
  recurrence: string | null;
  attendees: { name: string; response: string | null }[];
  body_markdown: string;
  truncated: boolean;
};

export interface CalendarSearch {
  /** YYYY-MM-DD. Defaults to the day Outlook's calendar opens. */
  start?: string;
  /** YYYY-MM-DD, inclusive. Defaults to start. */
  end?: string;
  /** Case-insensitive substring of the rendered title or location. Not Outlook search syntax. */
  query?: string;
  limit?: number;
}

// English headings only. A data-date attribute wins when Outlook provides one.
export function parseCalendarDate(text: string) {
  const cleaned = text.replace(/,/g, " ").replace(/\s+/g, " ").trim();
  let match = cleaned.match(/\b(\d{1,2}) ([A-Za-z]+) (\d{4})\b/);
  if (match) return calendarIso(Number(match[3]), match[2], Number(match[1]));
  match = cleaned.match(/\b([A-Za-z]+) (\d{1,2}) (\d{4})\b/);
  if (match) return calendarIso(Number(match[3]), match[1], Number(match[2]));
  return null;
}

function calendarIso(year: number, monthName: string, day: number) {
  const month = months.indexOf(monthName.toLowerCase());
  if (month < 0) return null;
  const date = new Date(Date.UTC(year, month, day));
  if (date.getUTCFullYear() !== year || date.getUTCMonth() !== month || date.getUTCDate() !== day) return null;
  return date.toISOString().slice(0, 10);
}

function assertIsoDate(value: string) {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value);
  if (!match) throw new MailError("INVALID_RANGE", "Dates must be YYYY-MM-DD.");
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const date = new Date(Date.UTC(year, month - 1, day));
  if (date.getUTCFullYear() !== year || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) {
    throw new MailError("INVALID_RANGE", "Dates must be real YYYY-MM-DD calendar dates.");
  }
}

function assertRange(start: string, end: string) {
  assertIsoDate(start);
  assertIsoDate(end);
  if (end < start) throw new MailError("INVALID_RANGE", "end is before start.");
  const span = Math.round((Date.parse(end + "T00:00:00Z") - Date.parse(start + "T00:00:00Z")) / 86_400_000);
  if (span > maxSpanDays) throw new MailError("INVALID_RANGE", "Calendar ranges are limited to 14 days.");
}

function eachDate(start: string, end: string) {
  const dates: string[] = [];
  const [year, month, day] = start.split("-").map(Number);
  const cursor = new Date(Date.UTC(year, month - 1, day));
  while (dates.length <= maxSpanDays) {
    const iso = cursor.toISOString().slice(0, 10);
    dates.push(iso);
    if (iso === end) return dates;
    cursor.setUTCDate(cursor.getUTCDate() + 1);
  }
  throw new MailError("INVALID_RANGE", "Calendar ranges are limited to 14 days.");
}

function clip(value: string, max: number) {
  const cleaned = value.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "").replace(/\s+/g, " ").trim();
  return Array.from(cleaned).slice(0, max).join("");
}

function settleTime(value: string, date: string, allDay: boolean) {
  const text = value.replace(/\s+/g, " ").trim();
  if (!text) return allDay ? date : "";
  if (/^\d{4}-\d{2}-\d{2}(?:[T\s]|$)/.test(text)) return text;
  return date + " " + text;
}

function sortKey(event: ListedEvent) {
  return [event.date, event.allDay ? "0" : "1", event.start, event.title].join("\u0000");
}

// Same walk as the mail reading pane so event notes follow the same Markdown rules:
// images become alt text, hidden nodes drop out, and links contribute only their text.
function eventSnapshot(dialog: Element, s: typeof selectors) {
  const textOf = (selector: string) => {
    const node = dialog.querySelector<HTMLElement>(selector);
    if (!node) return null;
    const style = getComputedStyle(node);
    if (style.visibility === "hidden" || style.visibility === "collapse" || !node.getClientRects().length) return null;
    return node.innerText;
  };
  const attribute = (selector: string) => dialog.querySelector(selector)?.getAttribute("datetime") ?? null;
  const attendees = Array.from(dialog.querySelectorAll(s.attendee)).slice(0, 40).flatMap(node => {
    const name = (node.querySelector("[data-testid='attendee-name']")?.textContent
      ?? node.textContent?.split(",")[0] ?? "").replace(/\s+/g, " ").trim();
    const status = (node.querySelector("[data-testid='attendee-status']")?.textContent
      ?? (node.textContent?.includes(",") ? node.textContent.split(",").slice(1).join(",") : "")).replace(/\s+/g, " ").trim();
    return name ? [{ name: name.slice(0, 200), response: status ? status.slice(0, 80) : null }] : [];
  });
  const notes = dialog.querySelector(s.notes);
  return { when: textOf(s.when), where: textOf(s.where), recurrence: textOf(s.recurrence),
    startAttr: attribute(s.start), endAttr: attribute(s.end), attendees, parts: notes ? eventNotes(notes) : [] };

  function eventNotes(root: Element): BodyPart[] {
    const parts: BodyPart[] = [];
    let buffer = "";
    const flush = (kind: BodyPart["kind"]) => {
      if (buffer.trim()) parts.push({ kind, text: buffer });
      buffer = "";
    };
    const walk = (node: Node, kind: BodyPart["kind"]) => {
      for (const child of node.childNodes) {
        if (child.nodeType === Node.TEXT_NODE) {
          buffer += (child.textContent ?? "").replace(/\s+/g, " ");
          continue;
        }
        if (child instanceof HTMLImageElement) {
          const alt = child.alt.replace(/\s+/g, " ").trim();
          if (alt.length > 1 && !child.hidden) buffer += " (image: " + alt + ") ";
          continue;
        }
        if (!(child instanceof HTMLElement) || child.hidden || child.getAttribute("aria-hidden") === "true" ||
            ["SCRIPT", "STYLE", "HEAD", "TITLE", "TEMPLATE", "svg"].includes(child.tagName)) continue;
        const style = getComputedStyle(child);
        if (style.display === "none" || style.visibility === "hidden") continue;
        if (child.tagName === "BR") {
          buffer += "\n";
          continue;
        }
        if (child instanceof HTMLTableRowElement && child.cells.length > 1 && !child.querySelector("table")) {
          flush(kind);
          const cells = Array.from(child.cells, cell => cell.innerText.replace(/\s+/g, " ").trim()).filter(Boolean);
          if (cells.length) parts.push({ kind: "row", text: cells.join(" | ") });
          continue;
        }
        const block = !style.display.startsWith("inline");
        const inner = child.tagName === "LI" ? "bullet" : child.tagName === "BLOCKQUOTE" ? "quote"
          : child.tagName === "HR" ? "separator" : kind;
        if (block) flush(kind);
        if (inner === "separator" && kind !== "separator") parts.push({ kind: "separator", text: "" });
        else walk(child, inner);
        if (block) flush(inner);
      }
    };
    walk(root, "paragraph");
    flush("paragraph");
    return parts;
  }
}

/**
 * Read-only calendar on the signed-in Outlook browser. It never creates, edits,
 * deletes or responds to events. A single click opens the details peek; a double
 * click would open the editor, so clicks stay single. Response and Save buttons
 * are never clicked. If the editor opens anyway, it is closed and unsaved changes
 * are discarded.
 */
export class Calendar {
  constructor(private readonly outlook: Outlook) {}

  private calendar?: Page;
  private verifiedDay?: { page: Page; url: string };
  private nav = 0;
  private refs = new Map<string, EventRef>();

  private surface(page: Page) {
    return page.locator(selectors.surface).or(page.getByRole("main", { name: /^calendar$/i })).first();
  }

  private dayUrl(date?: string, origin = mailbox.origin) {
    const path = date ? "/calendar/view/day/" + date.replaceAll("-", "/") : "/calendar/view/day";
    // A unique query avoids the browser treating a repeat of the same day as a no-op,
    // so each day is read from a fresh document.
    return origin + path + "?_=" + (++this.nav);
  }

  private async pageFor(context: BrowserContext) {
    const current = this.calendar;
    if (current && !current.isClosed() && current.context() === context) {
      current.setDefaultTimeout(8000);
      return current;
    }
    const page = await context.newPage();
    page.setDefaultTimeout(8000);
    this.outlook.holdPage(page);
    page.once("crash", () => { void page.close().catch(() => {}); });
    this.calendar = page;
    return page;
  }

  private authRequired(): never {
    // The worker signs back in once and repeats the calendar request before reporting this.
    throw new SessionLostError("Microsoft needs you to sign in again. " + signInHelp);
  }

  private async waitForCalendar(page: Page) {
    const surface = this.surface(page);
    const winner = await waitForSurface(page, surface, 20000, () => this.outlook.checkpoint());
    if (winner === "login" || isMicrosoftLogin(page)) this.authRequired();
    if (winner !== "ready") {
      if (page.url().startsWith("chrome-error:")) {
        throw new MailError("NETWORK_ERROR", "Could not reach Outlook. Check the network connection and try again.");
      }
      if (isMicrosoftLogin(page) || !outlookOrigins.has(safeOrigin(page.url()))) this.authRequired();
      throw new MailError("UI_CHANGED", "Could not identify the Outlook calendar.");
    }
  }

  private async readShownDate(page: Page) {
    const columns = this.surface(page).locator('[data-app-section="Surface_Day"] [data-column-date]');
    if (await columns.count()) {
      const dates = await columns.evaluateAll(nodes => [...new Set(nodes.map(n => n.getAttribute("data-column-date")))]);
      if (dates.length === 1 && dates[0] && /^\d{4}-\d{2}-\d{2}$/.test(dates[0])) return dates[0];
    }
    const dated = this.surface(page).locator(selectors.date).first();
    if (await dated.count()) {
      const iso = await dated.getAttribute("data-date");
      if (iso && /^\d{4}-\d{2}-\d{2}$/.test(iso)) return iso;
      const parsed = parseCalendarDate(await dated.innerText().catch(() => ""));
      if (parsed) return parsed;
    }
    const heading = this.surface(page).getByRole("heading").first();
    const parsed = parseCalendarDate(await heading.innerText({ timeout: 1000 }).catch(() => ""));
    if (!parsed) throw new MailError("UI_CHANGED", "Could not identify the calendar date.");
    return parsed;
  }

  private async stepButton(page: Page, direction: "next" | "previous") {
    const surface = this.surface(page);
    const specific = surface.getByRole("button", { name: direction === "next" ? /next day/i : /previous day/i }).first();
    if (await specific.waitFor({ state: "visible", timeout: 3000 }).then(() => true, () => false)) return specific;
    const broad = surface.getByRole("button", { name: direction === "next" ? /^next$/i : /^previous$/i }).first();
    if (await broad.isVisible().catch(() => false)) return broad;
    return null;
  }

  /** Moves the open day view onto date. The URL is tried first; buttons and a date input are fallbacks. */
  private async alignDate(page: Page, date: string) {
    if (await this.readShownDate(page) === date) return;
    const input = this.surface(page).locator('input[type="date"]').first();
    if (await input.count() && await input.isVisible().catch(() => false)) {
      await input.fill(date);
      await input.dispatchEvent("change");
      if (await this.readShownDate(page).catch(() => null) === date) return;
    }
    for (let step = 0; step < 16; step++) {
      const shown = await this.readShownDate(page);
      if (shown === date) return;
      const button = await this.stepButton(page, shown < date ? "next" : "previous");
      if (!button) break;
      await button.click();
      for (let attempt = 0; attempt < 40; attempt++) {
        if (await this.readShownDate(page).catch(() => shown) !== shown) break;
        await page.waitForTimeout(100);
      }
    }
    if (await this.readShownDate(page) !== date) {
      throw new MailError("UI_CHANGED", "Could not open the requested calendar day.");
    }
  }

  private async openDay(page: Page, date?: string) {
    this.outlook.checkpoint();
    if (this.verifiedDay?.page === page && this.verifiedDay.url === page.url()
        && outlookOrigins.has(safeOrigin(page.url())) && await this.surface(page).isVisible().catch(() => false)) {
      const shown = await this.readShownDate(page).catch(() => null);
      const dayView = !await page.locator('[data-app-section="CalendarModule"]').count()
        || await page.locator('[data-app-section="Surface_Day"]').count() > 0;
      if (shown && dayView && (!date || shown === date)) return shown;
    }
    this.verifiedDay = undefined;
    const origin = outlookOrigins.has(safeOrigin(page.url())) ? safeOrigin(page.url()) : mailbox.origin;
    await page.goto(this.dayUrl(date, origin), { waitUntil: "domcontentloaded", timeout: 20000 });
    await this.waitForCalendar(page);
    if (await page.locator('[data-app-section="CalendarModule"]').count()) {
      if (!await page.locator('[data-app-section="Surface_Day"]').count()) {
        await page.getByRole("button", { name: "Day", exact: true }).click();
      }
      await page.locator('[data-app-section="Surface_Day"] [data-column-date]').first().waitFor({ state: "attached" });
    }
    if (date) await this.alignDate(page, date);
    const shown = await this.readShownDate(page);
    if (date && shown !== date) throw new MailError("UI_CHANGED", "Could not open the requested calendar day.");
    this.verifiedDay = { page, url: page.url() };
    return shown;
  }

  private async collect(page: Page, date: string): Promise<ListedEvent[]> {
    const grid = this.surface(page).locator('[role="grid"], [data-app-section="Surface_Day"] [data-app-section="calendar-view-0"]').first();
    const attached = await grid.waitFor({ state: "attached", timeout: 8000 }).then(() => true, () => false);
    if (!attached) throw new MailError("UI_CHANGED", "Could not identify the calendar day grid.");
    if (await page.locator('[data-app-section="Surface_Day"]').count()) {
      await page.locator('[aria-live="polite"]').getByText(/^Loaded \d+ events? in /).first()
        .waitFor({ state: "attached", timeout: 15000 });
    }
    const extracted = await this.surface(page).locator(selectors.event).evaluateAll((nodes, day: string) => {
      const weekdays = /^(monday|tuesday|wednesday|thursday|friday|saturday|sunday)$/i;
      const range = /(\d{1,2}:\d{2}\s*(?:[AP]M)?)(?:\s*(?:to|–|—|-)\s*(\d{1,2}:\d{2}\s*(?:[AP]M)?))/i;
      const field = (node: Element, selector: string) =>
        node.querySelector(selector)?.textContent?.replace(/\s+/g, " ").trim() ?? "";
      const fromLabel = (label: string) => {
        const parts = label.split(",").map(part => part.replace(/\s+/g, " ").trim()).filter(Boolean);
        const time = label.match(range);
        let location: string | null = null;
        for (const part of parts.slice(1)) {
          if (/^(?:monday|tuesday|wednesday|thursday|friday|saturday|sunday|january|february|march|april|may|june|july|august|september|october|november|december)\b/i.test(part) || /^\d{4}$/.test(part) || /^Categorized as /i.test(part)) continue;
          if (/\ball[- ]day\b/i.test(part) || range.test(part) || /^\d{1,2}:\d{2}\s*(?:[AP]M)?$/i.test(part) || /^by\s/i.test(part)) continue;
          if (/^(recurring|recurs|private|busy|free|tentative|organizer)$/i.test(part)) continue;
          location = part;
          break;
        }
        return {
          title: parts[0] ?? "",
          allDay: /\ball[- ]day\b/i.test(label),
          start: time?.[1]?.replace(/\s+/g, " ").trim() ?? "",
          end: time?.[2]?.replace(/\s+/g, " ").trim() ?? "",
          location,
        };
      };
      const looksLikeEvent = (label: string) => {
        if (!label || /\b(create|new event)\b/i.test(label)) return false;
        if (/^(previous|next|today)\b/i.test(label)) return false;
        const parsed = fromLabel(label);
        if (!parsed.title || weekdays.test(parsed.title) || /^\d/.test(parsed.title)) return false;
        return parsed.allDay || Boolean(parsed.start);
      };
      const strict = nodes.some(node => node.getAttribute("data-testid") === "calendar-event");
      const events: RawEvent[] = [];
      let unreadable = false;
      nodes.forEach((node, index) => {
        if (node.closest("[role='dialog']")) return;
        const isStrict = node.getAttribute("data-testid") === "calendar-event";
        if (strict && !isStrict) return;
        const label = node.getAttribute("aria-label") ?? "";
        const parsed = fromLabel(label);
        if (!isStrict && !looksLikeEvent(label)) return;
        const title = field(node, "[data-testid='event-title']") || parsed.title;
        if (!title) {
          if (isStrict) unreadable = true;
          return;
        }
        const startNode = node.querySelector("[data-testid='event-start']");
        const endNode = node.querySelector("[data-testid='event-end']");
        const sourceId = node.getAttribute("data-event-id");
        events.push({
          index,
          sourceId: sourceId || null,
          title,
          start: startNode?.getAttribute("datetime") || field(node, "[data-testid='event-start']") || parsed.start,
          end: endNode?.getAttribute("datetime") || field(node, "[data-testid='event-end']") || parsed.end,
          allDay: node.getAttribute("data-all-day") === "true" || parsed.allDay,
          location: field(node, "[data-testid='event-location']") || parsed.location,
        });
      });
      return { events, unreadable, day };
    }, date);
    if (extracted.unreadable) throw new MailError("UI_CHANGED", "Could not identify an event title on the calendar.");
    return extracted.events.flatMap(raw => {
      const title = clip(raw.title, 300);
      if (!title) return [];
      const allDay = raw.allDay;
      let start = settleTime(raw.start, date, allDay);
      let end = settleTime(raw.end, date, allDay);
      if (!start) start = date;
      if (!end) end = start;
      const location = clip(raw.location ?? "", 500) || null;
      const event: ListedEvent = {
        index: raw.index, sourceId: raw.sourceId, title, start, end, allDay, location, date,
        fingerprint: JSON.stringify([raw.sourceId, title, start, end, location, date]),
      };
      return [event];
    });
  }

  private remember(event: ListedEvent) {
    const event_id = randomUUID();
    this.refs.set(event_id, {
      date: event.date, title: event.title, start: event.start, end: event.end, allDay: event.allDay,
      location: event.location, fingerprint: event.fingerprint, expires: Date.now() + idTtlMs,
    });
    return {
      event_id, title: event.title, start: event.start, end: event.end, all_day: event.allDay,
      location: event.location, date: event.date,
    };
  }

  private prune() {
    for (const [id, ref] of this.refs) {
      if (ref.expires <= Date.now() || this.refs.size > 500) this.refs.delete(id);
    }
  }

  private reference(id: string) {
    const ref = this.refs.get(id);
    if (!ref || ref.expires <= Date.now()) {
      throw new MailError("UNKNOWN_EVENT", "Run search_events again to get a fresh event ID.");
    }
    return ref;
  }

  async search(input: CalendarSearch = {}, signal?: AbortSignal) {
    const limit = input.limit ?? 10;
    if (!Number.isInteger(limit) || limit < 1 || limit > 25) {
      throw new MailError("INVALID_RANGE", "limit must be an integer from 1 to 25.");
    }
    if (input.end && !input.start) throw new MailError("INVALID_RANGE", "Pass start when end is set.");
    let query: string | undefined;
    if (input.query !== undefined) {
      query = input.query.trim();
      if (!query || query.length > 300) throw new MailError("INVALID_RANGE", "query must be 1 to 300 characters.");
    }
    if (input.start) assertRange(input.start, input.end ?? input.start);
    return this.outlook.useBrowser(async context => {
      const page = await this.pageFor(context);
      const start = input.start ?? await this.openDay(page);
      const end = input.end ?? start;
      assertRange(start, end);
      const collected: ListedEvent[] = [];
      const seen = new Set<string>();
      const needle = query?.toLowerCase();
      let scannedEnd = start;
      for (const date of eachDate(start, end)) {
        await this.openDay(page, date);
        const daily = await this.collect(page, date);
        for (const event of daily) {
          if (needle && !event.title.toLowerCase().includes(needle) && !(event.location ?? "").toLowerCase().includes(needle)) continue;
          if (seen.has(event.fingerprint)) continue;
          seen.add(event.fingerprint);
          collected.push(event);
        }
        scannedEnd = date;
        if (collected.length >= limit) break;
      }
      collected.sort((a, b) => sortKey(a).localeCompare(sortKey(b)));
      const events = collected.slice(0, limit).map(event => this.remember(event));
      this.prune();
      return { events, scope: "rendered_day_view" as const, start, end, query: query ?? null,
        scanned_start: start, scanned_end: scannedEnd, stopped_early: scannedEnd < end,
        truncated: scannedEnd < end || collected.length > limit };
    }, signal);
  }

  /**
   * Opens the event peek. Single click only. If Outlook shows the editor (a title
   * box plus Save), Escape and Discard run and the read fails closed — Save is not clicked.
   */
  async read(id: string, signal?: AbortSignal): Promise<EventDetails> {
    const ref = this.reference(id);
    return this.outlook.useBrowser(context => this.readInContext(context, ref, id), signal);
  }

  readEvents(ids: string[], signal?: AbortSignal) {
    if (!ids.length || ids.length > 10) throw new MailError("INVALID_RANGE", "Pass 1 to 10 event IDs.");
    return this.outlook.useBrowser(async context => {
      const pending = ids.map((event_id, index) => ({ event_id, index, date: this.refs.get(event_id)?.date ?? "" }));
      pending.sort((a, b) => a.date.localeCompare(b.date));
      const results: { event_id: string; data?: EventDetails; error?: ReturnType<typeof errorInfo> }[] = new Array(ids.length);
      for (const item of pending) {
        this.outlook.checkpoint();
        try {
          results[item.index] = { event_id: item.event_id, data: await this.readInContext(context, this.reference(item.event_id), item.event_id) };
        } catch (error) {
          this.outlook.checkpoint();
          if (error instanceof SessionLostError) throw error;
          results[item.index] = { event_id: item.event_id, error: errorInfo(error) };
        }
      }
      return { results };
    }, signal);
  }

  private async readInContext(context: BrowserContext, ref: EventRef, id: string): Promise<EventDetails> {
      this.outlook.checkpoint();
      const page = await this.pageFor(context);
      await this.openDay(page, ref.date);
      const events = await this.collect(page, ref.date);
      const match = events.find(event => event.fingerprint === ref.fingerprint);
      if (!match) throw new MailError("EVENT_CHANGED", "The event moved or is no longer shown. Search again.");
      const card = this.surface(page).locator(selectors.event).nth(match.index);
      const titleNode = card.locator(selectors.title).first();
      const target = await titleNode.count() ? titleNode : card;
      this.outlook.checkpoint();
      let editor = false;
      try {
        await target.click({ clickCount: 1 });
        const dialog = page.locator(selectors.peek).filter({ visible: true }).last();
        const appeared = await dialog.waitFor({ state: "visible", timeout: 8000 }).then(() => true, () => false);
        if (!appeared || isMicrosoftLogin(page)) {
          if (isMicrosoftLogin(page)) this.authRequired();
          throw new MailError("UI_CHANGED", "Could not find the event details.");
        }
        if (await this.isEditor(dialog)) {
          editor = true;
          await this.abandonEditor(page);
        }
        this.outlook.checkpoint();
        return await this.extractDetails(dialog, ref, id);
      } finally {
        if (!editor) await page.keyboard.press("Escape").catch(() => {});
      }
  }

  private async isEditor(dialog: ReturnType<Page["locator"]>) {
    const save = dialog.getByRole("button", { name: /^save$/i }).first();
    const textbox = dialog.getByRole("textbox").first();
    return await save.isVisible().catch(() => false) && await textbox.isVisible().catch(() => false);
  }

  private async abandonEditor(page: Page): Promise<never> {
    await page.keyboard.press("Escape");
    const discard = page.getByRole("button", { name: /^(discard|don't save|don’t save)$/i }).first();
    const visible = await discard.waitFor({ state: "visible", timeout: 1000 }).then(() => true, () => false);
    if (visible) await discard.click();
    throw new MailError("UI_CHANGED", "Outlook opened the event editor instead of the details view.");
  }

  private async extractDetails(dialog: ReturnType<Page["locator"]>, ref: EventRef, id: string): Promise<EventDetails> {
    const heading = dialog.locator(selectors.heading).first();
    const title = clip((await heading.innerText({ timeout: 2000 }).catch(() => "")).replace(/\s+/g, " ").trim(), 300);
    if (!title) throw new MailError("UI_CHANGED", "Could not find the event title in the details view.");
    if (title !== ref.title) throw new MailError("EVENT_CHANGED", "The event details did not match the calendar item. Search again.");
    const snapshot = await dialog.evaluate(eventSnapshot, selectors);
    const { startAttr, endAttr, attendees, parts } = snapshot;
    const when = clip(snapshot.when ?? "", 500) || null;
    const recurrence = clip(snapshot.recurrence ?? "", 500) || null;
    const whereText = (clip(snapshot.where ?? "", 500) || null) ?? ref.location;
    return {
      event_id: id, title, start: startAttr || ref.start, end: endAttr || ref.end, all_day: ref.allDay,
      location: whereText, date: ref.date, when, recurrence, attendees, ...formatBody(parts),
    };
  }
}
