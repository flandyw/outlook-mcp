import type { BrowserContext, Locator, Page } from "playwright";
import { randomUUID } from "node:crypto";
import { readdir, readFile, rm, stat } from "node:fs/promises";
import { launch, mailbox, outlookOrigins, profileDir, ProfileLockedError, safeOrigin } from "./browser.js";
import { inspectPasswordFile, isMicrosoftLogin, PasswordFileError, reauthenticate, resumeSignIn } from "./reauth.js";
import { formatBody, type BodyPart } from "./format.js";

export class MailError extends Error {
  constructor(readonly code: string, message: string) {
    super(message);
  }
}

type Row = { index: number; sender: string; subject: string; preview: string; time: string; fingerprint: string; hasAttachments: boolean };
type Reference = { query: string; row: Row; expires: number };

// Keep selectors scoped and fail closed when the expected Outlook layout is missing.
const selectors = {
  results: '[role="listbox"] [role="option"], [role="grid"] [role="row"]',
  subject: '[data-testid="message-subject"], [data-testid="subject"], .lvHighlightSubjectClass, .TtcXM',
  sender: '[data-testid="message-sender"], [data-testid="sender"], .lvHighlightFromClass, .ESO13 span[title]',
  preview: '[data-testid="message-preview"], [data-testid="preview"], .lvHighlightBodyClass, .ASFJj',
  pane: '[data-testid="reading-pane"], [aria-label="Reading Pane"], [aria-label="Reading pane"]',
  heading: '[data-testid="message-subject"], [role="heading"], h1, h2',
  // Ordered by preference. The first group with a visible match wins, so a generic
  // fallback never shadows the real body (Outlook also marks the subject .allowTextSelection).
  body: ['[id^="UniqueMessageBody"], [role="document"][aria-label="Message body"], [data-testid="message-body"]', '[role="document"]'],
  attachment: '[role="listbox"][aria-label*="attachment" i] [role="option"], [data-testid="attachment-name"]',
  // The card's "More actions" menu holds Download (checked on outlook.cloud.microsoft, Nov 2026).
  attachmentMenu: 'button[aria-haspopup="true"], button[aria-label="More actions"], [role="button"][aria-haspopup="true"]',
  recipients: { to: '[id$="_TO"][role="heading"]', cc: '[id$="_CC"][role="heading"]' },
  // Item id identifies one message. Conversation id is the fallback and can be shared by several rows.
  keyAttributes: ["data-item-id", "data-convid"],
};

const temporaryIdMs = 30 * 60_000;
/** Upper bound for `search_mail`'s limit. Higher values only help when Outlook will render more rows. */
export const searchResultLimit = 40;

const hasKey = (row: Row) => Boolean(JSON.parse(row.fingerprint)[0]);

function rowKey(row: Row) {
  try {
    const key: unknown = JSON.parse(row.fingerprint)[0];
    return typeof key === "string" && key ? key : null;
  } catch {
    return null;
  }
}

// A stable id carries the query and the Outlook key so read_mail can search again after the
// in-memory cache expires or the process restarts. Keyless rows keep a random UUID instead.
function durableMessageId(query: string, row: Row) {
  const key = rowKey(row);
  if (!key || key.length > 500 || query.length < 1 || query.length > 300) return null;
  if (row.sender.length > 500 || row.subject.length > 500 || row.time.length > 200) return null;
  const encode = (preview: string) => {
    const token = "ow1." + Buffer.from(JSON.stringify({
      v: 1, q: query, k: key, s: row.sender, j: row.subject, t: row.time,
      p: preview, a: row.hasAttachments ? 1 : 0,
    })).toString("base64url");
    return token.length <= 4096 ? token : null;
  };
  return encode(row.preview.length > 2000 ? "" : row.preview) ?? encode("");
}

function durableReference(id: string): Reference | null {
  const token = /^ow1\.([A-Za-z0-9_-]+)$/.exec(id);
  if (!token) return null;
  let value: unknown;
  try {
    value = JSON.parse(Buffer.from(token[1], "base64url").toString());
  } catch {
    return null;
  }
  if (!value || typeof value !== "object") return null;
  const { v, q, k, s, j, t, p, a } = value as Record<string, unknown>;
  if (v !== 1 || typeof q !== "string" || typeof k !== "string" || typeof s !== "string" ||
      typeof j !== "string" || typeof t !== "string" || typeof p !== "string" || (a !== 0 && a !== 1)) return null;
  if (q.length < 1 || q.length > 300 || k.length < 1 || k.length > 500) return null;
  if (s.length > 500 || j.length > 500 || t.length > 200 || p.length > 2000) return null;
  return {
    query: q,
    expires: Number.POSITIVE_INFINITY,
    row: {
      index: 0, sender: s, subject: j, preview: p, time: t, hasAttachments: a === 1,
      fingerprint: JSON.stringify([k, s, j, t, p]),
    },
  };
}

function searchRegion(page: Page) {
  return page.getByRole("region", { name: /search results/i })
    .or(page.getByRole("main", { name: /search results/i }))
    .or(page.locator('[data-testid="search-results"], [aria-label="Search results"], #MailList'))
    .filter({ visible: true }).first();
}

export type AuthState = "unknown" | "signed_in" | "signed_out";

export interface OutlookOptions {
  /** Close an idle browser after this many milliseconds. 0 keeps it open. */
  idleMs?: number;
  /** Reset the browser when one mail operation runs longer than this. */
  operationMs?: number;
  /** Reload Outlook before the next operation once the page has been open this long. */
  reloadMs?: number;
  /** How long to wait for Outlook's search box after a navigation. */
  readyMs?: number;
  /** How long temporary message IDs stay valid. Stable Outlook IDs ignore this. */
  referenceTtlMs?: number;
  /** Opens a visible browser for interactive sign-in. */
  launchVisible?: (display: Record<string, string>) => Promise<BrowserContext>;
  /** Called when the sign-in state changes because of a mail operation. */
  onAuthChange?: (state: AuthState, detail?: { code: string; message: string; at: string }) => void;
}

const interactiveMode = () => process.env.OUTLOOK_INTERACTIVE === "1";
const signInHelp = "Run npm run login on the server host to sign in again.";
const sessionExpired = "Microsoft session expired. " + signInHelp;
const loginWindowOpen = "A login window is already open on the server host.";

export class Outlook {
  constructor(
    private readonly launchBrowser: () => Promise<BrowserContext> = () => launch(true),
    private readonly options: OutlookOptions = {},
  ) {}

  private auth: AuthState = "unknown";
  private busy?: "signing_in" | "logging_out";
  private lastError?: { code: string; message: string; at: string };
  private lastSuccessAt?: string;
  private lastSignInAttempt?: number;
  private lastSearchQuery?: string;
  private loadedAt = 0;
  private idleTimer?: NodeJS.Timeout;
  private context?: BrowserContext;
  private page?: Page;
  private queue = Promise.resolve();
  private references = new Map<string, Reference>();

  status() {
    return {
      auth: this.auth,
      browser: this.context ? "running" : "stopped",
      busy: this.busy ?? null,
      last_success_at: this.lastSuccessAt ?? null,
      last_error: this.lastError ?? null,
      profile: profileDir(),
      password_file: inspectPasswordFile(),
    };
  }

  private setAuth(state: AuthState, detail?: { code: string; message: string; at: string }) {
    if (state === this.auth) return;
    this.auth = state;
    this.options.onAuthChange?.(state, detail);
  }

  private busyError() {
    return this.busy === "logging_out"
      ? new MailError("SIGNED_OUT", "The mailbox is being signed out.")
      : new MailError("AUTH_REQUIRED", loginWindowOpen);
  }

  private async ensurePage() {
    if (!this.context) {
      const context = await this.launchBrowser();
      this.context = context;
      context.once("close", () => {
        if (this.context !== context) return;
        this.context = undefined;
        this.page = undefined;
      });
    }
    if (!this.page || this.page.isClosed()) {
      const page = this.context.pages().find(candidate => !candidate.isClosed()) ?? await this.context.newPage();
      page.once("crash", () => {
        console.error("outlook-mcp: the Outlook tab crashed. The next request opens a new one.");
        void page.close().catch(() => {});
      });
      this.page = page;
      this.lastSearchQuery = undefined;
      this.loadedAt = 0;
    }
    this.page.setDefaultTimeout(8000);
    return this.page;
  }

  private async shutdownBrowser() {
    clearTimeout(this.idleTimer);
    const context = this.context;
    this.context = undefined;
    this.page = undefined;
    this.lastSearchQuery = undefined;
    await context?.close().catch(() => {});
  }

  private scheduleIdleClose() {
    const idle = this.options.idleMs ?? 0;
    clearTimeout(this.idleTimer);
    if (!idle || !this.context) return;
    this.idleTimer = setTimeout(() => {
      this.queue = this.queue.then(() => this.busy ? undefined : this.shutdownBrowser());
    }, idle);
    this.idleTimer.unref();
  }

  private classify(error: unknown) {
    if (error instanceof MailError) return error;
    if (error instanceof ProfileLockedError) return new MailError("PROFILE_LOCKED", error.message);
    const message = String(error instanceof Error ? error.message : error).split("\n")[0];
    // Local diagnostics only. The first line of a Playwright error names the failed step.
    console.error("outlook-mcp: " + message);
    if (this.page && !this.page.isClosed() && isMicrosoftLogin(this.page)) {
      return new MailError("AUTH_REQUIRED", sessionExpired);
    }
    if (/net::ERR_(?:INTERNET_DISCONNECTED|NAME_NOT_RESOLVED|NAME_RESOLUTION|CONNECTION_|NETWORK_|ADDRESS_UNREACHABLE|PROXY_|TIMED_OUT)/.test(message)) {
      return new MailError("NETWORK_ERROR", "Could not reach Outlook. Check the network connection and try again.");
    }
    if (error instanceof Error && error.name === "TimeoutError") {
      return new MailError("TIMEOUT", "Outlook took too long to respond. Try again.");
    }
    if (/has been closed|browser.*(?:closed|disconnected)|Target crashed/i.test(message)) {
      return new MailError("BROWSER_RESTARTED", "The browser closed unexpectedly and will restart. Try again.");
    }
    return new MailError("BROWSER_ERROR", "Outlook did not respond as expected. Check the browser session and layout.");
  }

  private run<T>(work: (page: Page) => Promise<T>) {
    if (this.busy) return Promise.reject(this.busyError());
    const result = this.queue.then(async () => {
      if (this.busy) throw this.busyError();
      clearTimeout(this.idleTimer);
      let watchdog: NodeJS.Timeout | undefined;
      try {
        const page = await this.ensurePage();
        const limit = this.options.operationMs ?? 90_000;
        const stuck = new Promise<never>((_, reject) => {
          watchdog = setTimeout(() => reject(new MailError("TIMEOUT", "Outlook stopped responding. The browser was reset; try again.")), limit);
        });
        const value = await Promise.race([work(page), stuck]);
        this.lastSuccessAt = new Date().toISOString();
        return value;
      } catch (error) {
        const failure = this.classify(error);
        this.lastError = { code: failure.code, message: failure.message, at: new Date().toISOString() };
        if (failure.code === "AUTH_REQUIRED") this.setAuth("signed_out", this.lastError);
        // A hung or crashed browser cannot be trusted with the next request.
        if (failure.code === "BROWSER_RESTARTED" || (failure.code === "TIMEOUT" && error === failure)) await this.shutdownBrowser();
        throw failure;
      } finally {
        clearTimeout(watchdog);
        this.scheduleIdleClose();
      }
    });
    this.queue = result.then(() => {}, () => {});
    return result;
  }

  async openSignIn() {
    return this.run(async page => {
      await page.goto(mailbox.href, { waitUntil: "domcontentloaded", timeout: 20000 });
      this.loadedAt = Date.now();
    });
  }

  /** Loads Outlook and reports whether the mailbox is usable. */
  async check() {
    await this.run(page => this.ensureMailbox(page));
    return this.status();
  }

  /**
   * Opens a visible browser window on the given display and waits until Outlook loads, the window
   * closes or the timeout passes. Mail requests fail fast with AUTH_REQUIRED meanwhile.
   */
  async signIn({ display = {}, timeoutMs = 10 * 60_000 }: { display?: Record<string, string>; timeoutMs?: number } = {}) {
    if (this.busy) throw new MailError("BUSY", "A sign-in or sign-out is already running.");
    this.busy = "signing_in";
    let context: BrowserContext | undefined;
    try {
      await this.queue;
      await this.shutdownBrowser();
      context = await (this.options.launchVisible ?? (env => launch(false, env)))(display);
      let closed = false;
      context.once("close", () => { closed = true; });
      const page = context.pages()[0] ?? await context.newPage();
      await page.goto(mailbox.href, { waitUntil: "domcontentloaded", timeout: 30000 }).catch(() => {});
      const deadline = Date.now() + timeoutMs;
      while (!closed && !page.isClosed() && Date.now() < deadline) {
        const ready = outlookOrigins.has(safeOrigin(page.url())) && await page.locator("#topSearchInput")
          .or(page.getByRole("searchbox")).first().isVisible().catch(() => false);
        if (ready) {
          // Give Outlook a moment to persist its tokens before the window closes.
          await page.waitForTimeout(3000).catch(() => {});
          this.lastSignInAttempt = undefined;
          this.lastError = undefined;
          this.setAuth("signed_in");
          return "signed_in" as const;
        }
        await new Promise(resolve => setTimeout(resolve, 1000));
      }
      return closed || page.isClosed() ? "cancelled" as const : "timeout" as const;
    } catch (error) {
      throw this.classify(error);
    } finally {
      this.busy = undefined;
      await context?.close().catch(() => {});
    }
  }

  /** Closes the browser and deletes the profile, which removes every saved Microsoft sign-in. */
  async logout() {
    if (this.busy) throw new MailError("BUSY", "A sign-in or sign-out is already running.");
    this.busy = "logging_out";
    try {
      await this.queue;
      await this.shutdownBrowser();
      const dir = profileDir();
      const entries = await readdir(dir).catch(() => null);
      if (entries?.length && !entries.includes("Local State") && !entries.includes("Default")) {
        throw new MailError("UNSAFE_PROFILE", "Refusing to delete " + dir + " because it does not look like a browser profile.");
      }
      await rm(dir, { recursive: true, force: true });
      this.references.clear();
      this.lastError = undefined;
      this.lastSuccessAt = undefined;
      this.auth = "signed_out";
      return { profile_deleted: dir };
    } finally {
      this.busy = undefined;
    }
  }

  private async recoverSession(page: Page) {
    if (!isMicrosoftLogin(page)) return;
    const username = process.env.OUTLOOK_USERNAME;
    // A remembered account usually signs in again without a secret.
    const resumed = await resumeSignIn(page, username).catch(() => "needs_user" as const);
    if (resumed === "signed_in") return;
    const report = inspectPasswordFile();
    if (report.configured && !report.valid) {
      throw new MailError("AUTH_REQUIRED", report.message ?? "Password file misconfigured.");
    }
    if (!username || !report.configured || !report.path) {
      throw new MailError("AUTH_REQUIRED", sessionExpired);
    }
    if (this.lastSignInAttempt !== undefined && Date.now() - this.lastSignInAttempt < 5 * 60_000) {
      throw new MailError("AUTH_REQUIRED", "Automatic sign-in is cooling down after a recent attempt. " + signInHelp);
    }
    this.lastSignInAttempt = Date.now();
    try {
      await reauthenticate(page, username, report.path);
      this.lastSignInAttempt = undefined;
    } catch (error) {
      // Browser errors can echo the password that was filled, so they are not forwarded.
      if (error instanceof PasswordFileError) {
        this.lastSignInAttempt = undefined;
        throw new MailError("AUTH_REQUIRED", error.message);
      }
      throw new MailError("AUTH_REQUIRED", "Automatic sign-in failed. " + signInHelp);
    }
  }

  private async signInPrompt(page: Page) {
    return page.getByRole("link", { name: /^sign in$/i }).or(page.getByRole("button", { name: /^sign in$/i }))
      .or(page.getByText(/session (?:has )?expired|signed out|sign in again/i)).first().isVisible().catch(() => false);
  }

  /** Navigates to Outlook when needed, recovers sign-in and returns the search box. */
  private async ensureMailbox(page: Page) {
    let blockedNavigation = false;
    const watch = (request: import("playwright").Request) => {
      if (request.isNavigationRequest() && !outlookOrigins.has(safeOrigin(request.url()))) blockedNavigation = true;
    };
    const ignoreBlocked = (error: unknown) => {
      if (!blockedNavigation) throw error;
    };
    page.on("request", watch);
    try {
      await this.recoverSession(page);
      if (page.url() !== "about:blank" && interactiveMode() && !outlookOrigins.has(safeOrigin(page.url()))) {
        throw new MailError("AUTH_REQUIRED", loginWindowOpen + " Complete sign-in in that window.");
      }
      const stale = this.loadedAt > 0 && Date.now() - this.loadedAt > (this.options.reloadMs ?? 6 * 3600_000);
      if (!outlookOrigins.has(safeOrigin(page.url())) || stale) {
        await page.goto(mailbox.href, { waitUntil: "domcontentloaded", timeout: 20000 }).catch(ignoreBlocked);
        this.loadedAt = Date.now();
        this.lastSearchQuery = undefined;
      }
      await this.recoverSession(page);
      const modernSearch = page.locator("#topSearchInput");
      const search = modernSearch.or(page.getByRole("searchbox")).or(page.getByRole("textbox", { name: "Search", exact: true }))
        .or(page.locator('input[placeholder="Search"]')).first();
      const appears = (timeout: number) => search.waitFor({ timeout }).then(() => true, () => false);
      const readyMs = this.options.readyMs ?? 20000;
      let ready = await appears(readyMs);
      if (!ready && isMicrosoftLogin(page)) {
        await this.recoverSession(page);
        ready = await appears(readyMs);
      } else if (!ready && outlookOrigins.has(safeOrigin(page.url())) && !blockedNavigation) {
        // Outlook occasionally shows an error, update or session page. One reload usually clears it.
        await page.reload({ waitUntil: "domcontentloaded", timeout: 20000 }).catch(ignoreBlocked);
        this.loadedAt = Date.now();
        this.lastSearchQuery = undefined;
        await this.recoverSession(page);
        ready = await appears(readyMs);
      }
      if (!ready) {
        if (page.url().startsWith("chrome-error:") && !blockedNavigation) {
          throw new MailError("NETWORK_ERROR", "Could not reach Outlook. Check the network connection and try again.");
        }
        if (isMicrosoftLogin(page) || blockedNavigation || await this.signInPrompt(page)) {
          throw new MailError("AUTH_REQUIRED", sessionExpired);
        }
        throw new MailError("UI_CHANGED", "Could not find the Outlook search box.");
      }
      this.setAuth("signed_in");
      return { search, modernSearch };
    } finally {
      page.off("request", watch);
    }
  }

  private resultList(page: Page) {
    return searchRegion(page).locator(selectors.results).filter({ visible: true });
  }

  private async rows(page: Page, query: string): Promise<Row[]> {
    const existing = this.lastSearchQuery === query && await this.showingSearch(page, query)
      ? await this.extractRows(this.resultList(page))
      : [];
    if (existing.length) return existing;
    this.lastSearchQuery = undefined;
    const { search, modernSearch } = await this.ensureMailbox(page);
    const modern = await modernSearch.isVisible();
    if (modern) {
      await modernSearch.click();
      await modernSearch.press("Control+a");
      await modernSearch.pressSequentially(query, { delay: 20 });
    } else {
      await search.fill(query);
    }
    const oldRow = modern
      ? await page.locator('#MailList [role="option"]:visible').first().elementHandle({ timeout: 500 }).catch(() => null)
      : null;
    const oldRowId = oldRow ? await oldRow.getAttribute("id") : null;
    const searchResponse = modern ? page.waitForResponse(response =>
      new URL(response.url()).pathname === "/searchservice/api/v2/query" &&
      response.request().method() === "POST" && response.ok(), { timeout: 10000 })
      .then(() => true, () => false) : undefined;
    await (modern ? modernSearch : search).press("Enter");
    if (oldRow) {
      try {
        await page.waitForFunction(({ row, id }) =>
          !row.isConnected || !row.getClientRects().length || row.id !== id,
        { row: oldRow, id: oldRowId }, { timeout: 8000 });
      } catch {
        throw new MailError("UI_CHANGED", "Outlook did not replace the inbox rows with search results.");
      } finally {
        await oldRow.dispose();
      }
    }
    const candidates = this.resultList(page);
    const empty = page.getByText(/^(?:no results(?: found)?|we (?:couldn.t|didn.t) find anything)[.!]?$/i)
      .filter({ visible: true }).first();
    if (searchResponse && !await searchResponse) {
      throw new MailError("UI_CHANGED", "Outlook did not complete the search request.");
    }
    await candidates.first().or(empty).first().waitFor().catch(() => {
      throw new MailError("UI_CHANGED", "Could not identify the Outlook search results.");
    });
    if (await empty.isVisible()) return [];
    const rows = await this.extractRows(candidates);
    if (!rows.length) throw new MailError("UI_CHANGED", "Could not identify sender and subject in search results.");
    this.lastSearchQuery = query;
    return rows;
  }

  // Cached rows are reused only while the search box still holds the query, so the
  // inbox list is never mistaken for search results after someone clears the search.
  private async showingSearch(page: Page, query: string) {
    if (!outlookOrigins.has(safeOrigin(page.url()))) return false;
    const box = page.locator("#topSearchInput").or(page.getByRole("searchbox")).first();
    const value = await box.inputValue({ timeout: 500 }).catch(() => null);
    return value === null || value.trim() === query.trim();
  }

  private extractRows(candidates: Locator) {
    return candidates.evaluateAll((nodes, s) => nodes.flatMap((node, index) => {
      const text = (selector: string) => node.querySelector(selector)?.textContent?.trim() ?? "";
      const subject = text(s.subject);
      const sender = text(s.sender);
      if (!subject || !sender) return [];
      const timeNode = node.querySelector("time");
      const time = timeNode?.getAttribute("datetime") ?? timeNode?.textContent?.trim() ?? node.querySelector(".qq2gS[title]")?.getAttribute("title") ?? "";
      const preview = text(s.preview);
      let key: string | null = null;
      for (const name of s.keyAttributes) {
        const own = node.getAttribute(name);
        const nested = own || node.querySelector("[" + name + "]")?.getAttribute(name);
        if (nested) { key = nested; break; }
      }
      const fingerprint = JSON.stringify([key, sender, subject, time, preview]);
      // Outlook's row label says "Has attachments" (or similar) for messages with files.
      const hasAttachments = /\battach/i.test(node.getAttribute("aria-label") ?? "");
      return [{ index, sender, subject, preview, time, fingerprint, hasAttachments }];
    }), selectors);
  }

  // Outlook can list one message twice (top results and all results). Identical rows that carry an
  // Outlook item key are the same message, so keep the first. Keyless look-alikes stay ambiguous.
  private hitsFrom(rows: Row[]) {
    return rows.filter((row, position) => {
      const same = rows.filter(other => other.fingerprint === row.fingerprint);
      return same.length === 1 || (hasKey(row) && same[0] === rows[position]);
    });
  }

  /**
   * Collect identifiable rows from the top of the result list until `limit` hits are available or
   * Outlook stops rendering more. Virtualized lists only keep a window of rows in the DOM.
   */
  private async gatherRows(page: Page, query: string, limit: number) {
    const seeded = await this.rows(page, query);
    if (!seeded.length) return seeded;
    const collected: Row[] = [];
    const seen = new Set<string>();
    const add = (list: Row[]) => {
      // Keep every copy from this window so duplicate top/all-results rows still collapse later,
      // but don't count the same row again after the next scroll.
      const fresh = list.filter(row => !seen.has(row.fingerprint));
      collected.push(...fresh);
      for (const row of list) seen.add(row.fingerprint);
    };
    if (await this.scrollList(page, "top")) await page.waitForTimeout(250);
    add(await this.extractRows(this.resultList(page)));
    let stuck = 0;
    for (let step = 0; step < 50 && this.hitsFrom(collected).length < limit; step++) {
      const moved = await this.scrollList(page, "next");
      if (!moved) break;
      await page.waitForTimeout(250);
      const before = collected.length;
      add(await this.extractRows(this.resultList(page)));
      if (collected.length === before) {
        if (++stuck >= 3) break;
      } else stuck = 0;
    }
    return collected;
  }

  search(query: string, limit: number) {
    return this.run(async page => {
      const rows = await this.gatherRows(page, query, limit);
      const hits = this.hitsFrom(rows);
      const ttl = this.options.referenceTtlMs ?? temporaryIdMs;
      const messages = hits.slice(0, limit).map(row => {
        const message_id = durableMessageId(query, row) ?? randomUUID();
        this.references.set(message_id, { query, row, expires: Date.now() + ttl });
        return { message_id, sender: row.sender, subject: row.subject, received_at: row.time || null,
          has_attachments: row.hasAttachments, snippet: row.preview.slice(0, 300) };
      });
      for (const [id, reference] of this.references) {
        if (reference.expires <= Date.now() || this.references.size > 500) this.references.delete(id);
      }
      const omitted = rows.filter(row => !hits.some(hit => hit.fingerprint === row.fingerprint)).length;
      return { messages, scope: "rendered_search_results", ambiguous_results_omitted: omitted };
    });
  }

  private reference(id: string) {
    const cached = this.references.get(id);
    if (cached && cached.expires > Date.now()) return cached;
    const durable = durableReference(id);
    if (durable) return durable;
    if (cached) this.references.delete(id);
    throw new MailError("UNKNOWN_MESSAGE", "Run search_mail again to get a fresh message ID.");
  }

  /** Finds the search row for a message ID, verifies it and opens it in the reading pane. */
  private async openMessage(page: Page, reference: Reference) {
    if (reference.expires <= Date.now()) {
      throw new MailError("UNKNOWN_MESSAGE", "Run search_mail again to get a fresh message ID.");
    }
    const visibleRows = () => this.resultList(page);
    const sameMessage = (row: Row) => {
      if (row.fingerprint === reference.row.fingerprint) return true;
      const [key] = JSON.parse(row.fingerprint);
      const [expected] = JSON.parse(reference.row.fingerprint);
      return Boolean(key && key === expected && row.sender === reference.row.sender && row.subject === reference.row.subject &&
        row.time === reference.row.time);
    };
    // Exact fingerprints first. Several messages of one conversation share a key, sender and subject,
    // so the looser key match is only a fallback for rows whose preview Outlook re-rendered.
    const find = (rows: Row[]) => {
      const exact = rows.filter(row => row.fingerprint === reference.row.fingerprint);
      return exact.length ? exact : rows.filter(sameMessage);
    };
    let matches = this.lastSearchQuery === reference.query && await this.showingSearch(page, reference.query)
      ? await this.scrollToFind(page, find)
      : [];
    if (matches.length !== 1) {
      // Run the search again for real: rows() would otherwise hand back the same cached list.
      this.lastSearchQuery = undefined;
      const rows = await this.rows(page, reference.query);
      matches = find(rows);
      if (!matches.length) matches = await this.scrollToFind(page, find);
      if (!matches.length && reference.row.time) {
        matches = rows.filter(row => row.sender === reference.row.sender && row.subject === reference.row.subject &&
          row.time === reference.row.time && row.preview === reference.row.preview);
      }
    }
    // Outlook lists one message twice (top results and all results) and may re-render one copy's preview.
    // Copies with the same item key, sender, subject and time are the same message.
    const identity = (row: Row) => JSON.stringify([JSON.parse(row.fingerprint)[0], row.sender, row.subject, row.time]);
    if (matches.length > 1 && hasKey(matches[0]) && matches.every(match => identity(match) === identity(matches[0]))) {
      matches = [matches[0]];
    }
    if (matches.length !== 1) {
      console.error("outlook-mcp: message " + (matches.length ? "matched " + matches.length + " rows" : "not found") + " after a fresh search.");
      throw new MailError("MESSAGE_CHANGED", "The result moved or became ambiguous. Search again.");
    }
    const row = matches[0];
    const candidate = visibleRows().nth(row.index);
    const live = await this.extractRows(candidate);
    if (live[0]?.fingerprint !== row.fingerprint) {
      console.error("outlook-mcp: search row changed between matching and opening.");
      throw new MailError("MESSAGE_CHANGED", "Search results changed before opening the message.");
    }
    const wasUnread = await this.rowWasUnread(candidate);
    await candidate.locator(selectors.subject).first().click();
    return { row, wasUnread };
  }

  // Moves the list's scroll container: "top" resets it, "next" advances most of a screen. Returns false at the end.
  private scrollList(page: Page, step: "top" | "next") {
    return searchRegion(page).evaluate((root, step) => {
      const scrollable = [root, ...root.querySelectorAll("*")].find(element =>
        element.scrollHeight > element.clientHeight + 4 && /auto|scroll/.test(getComputedStyle(element).overflowY));
      if (!scrollable) return false;
      const before = scrollable.scrollTop;
      scrollable.scrollTop = step === "top" ? 0 : before + scrollable.clientHeight * 0.8;
      return step === "top" ? before !== 0 : scrollable.scrollTop !== before;
    }, step).catch(() => false);
  }

  /**
   * Outlook virtualizes the message list: only rows near the viewport exist in the page, and opening
   * messages scrolls the list. Look at the rendered rows, then page through the list from the top.
   */
  private async scrollToFind(page: Page, find: (rows: Row[]) => Row[]) {
    const visibleRows = () => this.resultList(page);
    let matches = find(await this.extractRows(visibleRows()));
    if (matches.length) return matches;
    let moved = await this.scrollList(page, "top");
    for (let steps = 0; steps < 25; steps++) {
      if (moved) await page.waitForTimeout(250);
      matches = find(await this.extractRows(visibleRows()));
      if (matches.length) return matches;
      moved = await this.scrollList(page, "next");
      if (!moved) break;
    }
    return matches;
  }

  // Unread is a status prefix or a dedicated badge, not a word that happens to appear in the subject.
  private rowWasUnread(row: Locator) {
    return row.evaluate(node => {
      if (/^\s*unread\b/i.test(node.getAttribute("aria-label") ?? "")) return true;
      return [...node.querySelectorAll("[aria-label]")].some(element =>
        /^\s*unread\s*$/i.test(element.getAttribute("aria-label") ?? ""));
    });
  }

  /**
   * After a message that was unread is opened, click Outlook's own Mark as unread control when exactly
   * one is visible outside the message body. Buttons inside the body are message content and are ignored.
   */
  private async restoreUnread(page: Page) {
    const click = () => page.evaluate(() => {
      const bodies = [...document.querySelectorAll('[id^="UniqueMessageBody"], [role="document"], [data-testid="message-body"]')];
      const named = [...document.querySelectorAll("button, [role='button']")].filter(element => {
        if (!(element instanceof HTMLElement) || element.hidden) return false;
        if (bodies.some(body => body.contains(element))) return false;
        const name = (element.getAttribute("aria-label") || element.innerText || "").replace(/\s+/g, " ").trim();
        if (!/^mark as unread$/i.test(name)) return false;
        const style = getComputedStyle(element);
        return style.display !== "none" && style.visibility !== "hidden" && element.getClientRects().length > 0;
      });
      if (named.length !== 1 || !(named[0] instanceof HTMLElement)) return false;
      named[0].click();
      return true;
    }).catch(() => false);
    const started = Date.now();
    while (Date.now() - started < 1000) {
      if (await click()) return;
      await page.waitForTimeout(100);
    }
  }

  private pane(page: Page, row: Row) {
    return page.locator(selectors.pane).filter({
      has: page.locator(selectors.heading).and(page.getByText(row.subject, { exact: true })),
    }).first();
  }

  /** One name per attachment card, in card order, so an index always selects the same card. */
  private attachmentNames(pane: Locator) {
    return pane.locator(selectors.attachment).evaluateAll(nodes => nodes.map((node, index) => {
      if (node.getAttribute("data-testid") === "attachment-name") return (node.textContent ?? "").trim() || "attachment-" + (index + 1);
      // File cards have a titled name. Image cards show a thumbnail, so fall back to the card's label,
      // which reads like "photo.jpg Open 2 MB".
      const titled = node.querySelector("[title]")?.getAttribute("title")?.trim();
      const text = ((node as HTMLElement).innerText ?? "").split("\n")[0].trim();
      const label = (node.getAttribute("aria-label") ?? "").replace(/\s+Open(?:\s+[\d.,]+\s*[KMGT]?B)?\s*$/i, "").trim();
      return titled || text || label || "attachment-" + (index + 1);
    }));
  }

  async read(id: string) {
    const reference = this.reference(id);
    return this.run(async page => {
      const opened = await this.openMessage(page, reference);
      try {
        return await this.extract(page, opened.row, id);
      } finally {
        if (opened.wasUnread) await this.restoreUnread(page);
      }
    });
  }

  /**
   * Downloads one attachment of a message through Outlook's own Download command and returns its bytes.
   * `attachment` is a zero-based index into attachment_names or an exact file name.
   */
  async downloadAttachment(id: string, attachment: number | string, maxBytes: number) {
    const reference = this.reference(id);
    return this.run(async page => {
      const opened = await this.openMessage(page, reference);
      const row = opened.row;
      try {
        const pane = this.pane(page, row);
        await pane.locator(selectors.attachment).first().waitFor({ timeout: 8000 }).catch(() => {
          throw new MailError("NO_ATTACHMENTS", "This message has no downloadable attachments.");
        });
        const names = await this.attachmentNames(pane);
        let index: number;
        if (typeof attachment === "number") {
          index = attachment;
        } else {
          let found = names.flatMap((name, position) => name === attachment ? [position] : []);
          if (!found.length) found = names.flatMap((name, position) => name.toLowerCase() === attachment.toLowerCase() ? [position] : []);
          if (found.length > 1) throw new MailError("AMBIGUOUS_ATTACHMENT", "Several attachments share that name. Pass its index instead.");
          index = found[0] ?? -1;
        }
        if (!(index >= 0 && index < names.length)) {
          throw new MailError("UNKNOWN_ATTACHMENT", "No such attachment. Available: " + JSON.stringify(names));
        }
        const card = pane.locator(selectors.attachment).nth(index);
        await card.hover().catch(() => {});
        const menu = card.locator(selectors.attachmentMenu).first();
        if (!await menu.count()) throw new MailError("ATTACHMENT_UNAVAILABLE", "Outlook offers no download for this attachment.");
        await menu.click();
        // Menu items carry icon glyphs in their text and often no aria-label, so match loosely but skip "Download all".
        const items = page.locator('[role="menuitem"]:visible');
        await items.first().waitFor({ timeout: 4000 }).catch(() => {});
        const labels = await items.evaluateAll(nodes => nodes.map(node =>
          ((node.getAttribute("aria-label") || node.textContent) ?? "").replace(/[^\p{L}\p{N} ]+/gu, " ").replace(/\s+/g, " ").trim()));
        const position = labels.findIndex(label => /^download$/i.test(label)) >= 0
          ? labels.findIndex(label => /^download$/i.test(label))
          : labels.findIndex(label => /\bdownload\b/i.test(label) && !/\ball\b/i.test(label));
        if (position < 0) {
          console.error("outlook-mcp: attachment menu had no Download item: " + JSON.stringify(labels));
          await page.keyboard.press("Escape").catch(() => {});
          throw new MailError("ATTACHMENT_UNAVAILABLE", "Outlook offers no download for this attachment (it may be a cloud link).");
        }
        const item = items.nth(position);
        const [download] = await Promise.all([page.waitForEvent("download", { timeout: 60_000 }), item.click()]);
        try {
          if (await download.failure()) throw new MailError("DOWNLOAD_FAILED", "Outlook could not download the attachment.");
          const path = await download.path();
          const { size } = await stat(path);
          if (size > maxBytes) {
            throw new MailError("ATTACHMENT_TOO_LARGE", "The attachment is " + Math.ceil(size / 1048576) + " MB, over the " +
              Math.floor(maxBytes / 1048576) + " MB limit.");
          }
          return { message_id: id, name: names[index] || download.suggestedFilename(), size_bytes: size, data: await readFile(path) };
        } finally {
          await download.delete().catch(() => {});
        }
      } finally {
        if (opened.wasUnread) await this.restoreUnread(page);
      }
    });
  }

  private async extract(page: Page, row: Row, id: string) {
    const pane = this.pane(page, row);
    const findBody = async (visible: boolean) => {
      for (const selector of selectors.body) {
        const candidate = visible ? pane.locator(selector).filter({ visible: true }) : pane.locator(selector);
        if (await candidate.count()) return candidate;
      }
    };
    let bodies: Locator | undefined;
    const started = Date.now();
    while (!bodies && Date.now() - started < 8000) {
      bodies = await findBody(true);
      // An empty or image-only body has no size, so Playwright never counts it as visible.
      if (!bodies && Date.now() - started > 2000) bodies = await findBody(false);
      if (!bodies) await page.waitForTimeout(150);
    }
    if (!bodies) throw new MailError("UI_CHANGED", "Could not find the message body in the reading pane.");
    // Outermost bodies only. A conversation shows one body per expanded message.
    const parts = await bodies.evaluateAll(nodes => {
      const roots = nodes.filter(node => !nodes.some(other => other !== node && other.contains(node)));
      const parts: BodyPart[] = [];
      roots.forEach((root, index) => {
        if (index) parts.push({ kind: "separator", text: "" });
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
            // Data tables become "a | b" rows. Layout tables (one cell or nested tables) are walked normally.
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
      });
      return parts;
    });
    // Attachment cards can render after the body. Wait for them only when the row says there are some.
    if (row.hasAttachments) await pane.locator(selectors.attachment).first().waitFor({ timeout: 8000 }).catch(() => {});
    const attachments = await this.attachmentNames(pane);
    const recipients = async (selector: string) => {
      const text = await pane.locator(selector).first().textContent({ timeout: 250 }).catch(() => null);
      return text?.replace(/^\s*(?:to|cc)\s*:\s*/i, "").replace(/\s+/g, " ").trim() || null;
    };
    return { message_id: id, subject: row.subject, sender: row.sender, received_at: row.time || null,
      to: await recipients(selectors.recipients.to), cc: await recipients(selectors.recipients.cc),
      attachment_names: attachments, ...formatBody(parts) };
  }

  async close() {
    await this.shutdownBrowser();
  }
}
