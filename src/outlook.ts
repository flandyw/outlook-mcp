import type { BrowserContext, Locator, Page } from "playwright";
import { randomUUID } from "node:crypto";
import { AsyncLocalStorage } from "node:async_hooks";
import { readdir, rm, stat, open } from "node:fs/promises";
import { launch, mailbox, outlookOrigins, profileDir, ProfileLockedError, safeOrigin, signInOrigins } from "./browser.js";
import { credentialsSummary, inspectPasswordFile, isMicrosoftLogin, PasswordFileError, passwordSource, reauthenticate, resumeSignIn, SignInError } from "./reauth.js";
import { formatBody, type BodyPart } from "./format.js";

export class MailError extends Error {
  constructor(readonly code: string, message: string) {
    super(message);
  }
}

export function errorInfo(error: unknown) {
  const code = error instanceof MailError ? error.code : "INTERNAL_ERROR";
  return { code, message: error instanceof MailError ? error.message : "The operation failed.",
    retryable: ["TIMEOUT", "NETWORK_ERROR", "BROWSER_RESTARTED", "DOWNLOAD_FAILED", "STORAGE_FULL"].includes(code) };
}

/**
 * Thrown when Microsoft signs the browser out partway through an operation (for example in the calendar tab).
 * The worker recovers the session once and repeats the operation before reporting AUTH_REQUIRED.
 */
export class SessionLostError extends MailError {
  constructor(message = "Microsoft signed the browser out. Run npm run login on the server host to sign in again.") {
    super("AUTH_REQUIRED", message);
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

function mailMetadata(root: Element, { s, row }: { s: typeof selectors; row: Row }) {
  const recipients = (selector: string) => root.querySelector(selector)?.textContent
    ?.replace(/^\s*(?:to|cc)\s*:\s*/i, "").replace(/\s+/g, " ").trim() || null;
  const cards = Array.from(root.querySelectorAll(s.attachment));
  const attachment_names = cards.map((node, index) => {
    if (node.getAttribute("data-testid") === "attachment-name") return node.textContent?.trim() || "attachment-" + (index + 1);
    const titled = node.querySelector("[title]")?.getAttribute("title")?.trim();
    const text = ((node as HTMLElement).innerText ?? "").split("\n")[0].trim();
    const label = (node.getAttribute("aria-label") ?? "").replace(/\s+Open(?:\s+[\d.,]+\s*[KMGT]?B)?\s*$/i, "").trim();
    return titled || text || label || "attachment-" + (index + 1);
  });
  const attachments = cards.map((node, index) => {
    const name = attachment_names[index];
    const sizePattern = /^(?:Size:\s*)?([\d.,]+)\s*(KB|MB|GB|TB|B)$/i;
    const filename = Array.from(node.querySelectorAll("*")).find(element =>
      element.getAttribute("title")?.trim() === name || element.textContent?.trim() === name);
    const sizeText = [node.querySelector('[data-testid="attachment-size"]'), ...node.querySelectorAll("*")]
      .filter((element): element is Element => !!element && !filename?.contains(element))
      .map(element => element.textContent?.trim() ?? "")
      .find(text => text !== name && sizePattern.test(text));
    const label = node.getAttribute("aria-label") ?? "";
    const suffix = label.startsWith(name) ? label.slice(name.length).replace(/^\s+Open\s+/i, "").trim() : "";
    const match = (sizeText ?? suffix).match(sizePattern);
    const size_bytes = match ? Math.round(Number(match[1].replace(/,/g, "")) * 1024 ** ["B", "KB", "MB", "GB", "TB"].indexOf(match[2].toUpperCase())) : null;
    const typeNode = node.querySelector('[data-testid="attachment-type"], img[alt]');
    const type = typeNode?.textContent?.trim() || typeNode?.getAttribute("alt")?.trim() || null;
    return { index, name, size_bytes, type };
  });
  return { subject: row.subject, sender: row.sender, to: recipients(s.recipients.to), cc: recipients(s.recipients.cc), attachment_names, attachments };
}

const temporaryIdMs = 30 * 60_000;
/** Upper bound for `search_mail`'s limit. Higher values only help when Outlook will render more rows. */
export const searchResultLimit = 40;

export async function waitForSurface(page: Page, surface: Locator, timeout: number, checkpoint = () => {}) {
  const deadline = Date.now() + timeout;
  do {
    checkpoint();
    if (isMicrosoftLogin(page)) return "login" as const;
    if (await surface.isVisible().catch(() => false)) return "ready" as const;
    if (page.isClosed()) break;
    const remaining = deadline - Date.now();
    if (remaining <= 0) break;
    await new Promise(resolve => setTimeout(resolve, Math.min(50, remaining)));
  } while (Date.now() < deadline);
  return "timeout" as const;
}

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
  /** Cooldown after a failed automatic sign-in. Rejected credentials wait longer to avoid lockouts. */
  private signInBlockedUntil = 0;
  private signInBlockedReason?: string;
  /** Set when the current operation already tried to recover the session. */
  private recoveryTried = false;
  private lastSearchQuery?: string;
  private loadedAt = 0;
  private idleTimer?: NodeJS.Timeout;
  private context?: BrowserContext;
  private page?: Page;
  /** Tabs owned by another read-only surface, such as calendar. Mail must not type into them. */
  private reserved = new Set<Page>();
  private queue = Promise.resolve();
  private operation = new AsyncLocalStorage<() => void>();
  private references = new Map<string, Reference>();

  checkpoint() {
    this.operation.getStore()?.();
  }

  status() {
    return {
      auth: this.auth,
      browser: this.context ? "running" : "stopped",
      busy: this.busy ?? null,
      last_success_at: this.lastSuccessAt ?? null,
      last_error: this.lastError ?? null,
      profile: profileDir(),
      password_file: inspectPasswordFile(),
      credentials: credentialsSummary(),
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
      try {
        this.checkpoint();
      } catch (error) {
        await context.close().catch(() => {});
        throw error;
      }
      this.context = context;
      context.once("close", () => {
        if (this.context !== context) return;
        this.context = undefined;
        this.page = undefined;
        this.reserved.clear();
      });
    }
    if (!this.page || this.page.isClosed()) {
      // Skip tabs another surface reserved. An empty set keeps the original "first open page" behavior.
      const page = this.context.pages().find(candidate => !candidate.isClosed() && !this.reserved.has(candidate))
        ?? await this.context.newPage();
      this.checkpoint();
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
    this.reserved.clear();
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

  private run<T>(work: (page: Page) => Promise<T>, signal?: AbortSignal) {
    if (this.busy) return Promise.reject(this.busyError());
    const deadline = Date.now() + (this.options.operationMs ?? 180_000);
    const timeout = new MailError("TIMEOUT", "Outlook operation exceeded its deadline. Try again.");
    const cancelled = new MailError("CANCELLED", "The request was cancelled.");
    let expired = false;
    let running = false;
    let abort: () => void;
    const aborted = new Promise<never>((_, reject) => { abort = () => reject(cancelled); });
    const onAbort = () => abort();
    if (signal?.aborted) onAbort();
    else signal?.addEventListener("abort", onAbort, { once: true });
    let watchdog: NodeJS.Timeout;
    const stuck = new Promise<never>((_, reject) => {
      watchdog = setTimeout(() => { expired = true; reject(timeout); }, Math.max(0, deadline - Date.now()));
    });
    const checkpoint = () => {
      if (expired || Date.now() >= deadline) throw timeout;
      if (signal?.aborted) throw cancelled;
    };
    const result = this.queue.then(async () => {
      checkpoint();
      running = true;
      if (this.busy) throw this.busyError();
      clearTimeout(this.idleTimer);
      return this.operation.run(checkpoint, async () => {
      try {
        this.recoveryTried = false;
        const attempt = async () => {
          const page = await this.ensurePage();
          checkpoint();
          try {
            const value = await work(page);
            checkpoint();
            return value;
          } catch (error) {
            checkpoint();
            if (!await this.sessionLost(page, error)) throw error;
            // Microsoft signed the browser out mid-operation. Recover once, then repeat the work.
            console.error("outlook-mcp: the Microsoft session ended during a request; signing back in.");
            await this.ensureMailbox(page, true);
            checkpoint();
            return await work(page);
          }
        };
        const value = await Promise.race([attempt(), stuck]);
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
        this.scheduleIdleClose();
      }
      });
    }).finally(() => {
      clearTimeout(watchdog);
      signal?.removeEventListener("abort", onAbort);
    });
    this.queue = result.then(() => {}, () => {});
    const pendingTimeout = stuck.catch(error => running ? new Promise<never>(() => {}) : Promise.reject(error));
    return Promise.race([result, pendingTimeout, aborted]);
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

  /** Explicit unattended login bypasses the automatic retry cooldown. */
  async signInNonInteractive() {
    this.clearSignInCooldown();
    try {
      await this.check();
      this.lastError = undefined;
      return "signed_in" as const;
    } catch (error) {
      if (error instanceof MailError && error.code === "AUTH_REQUIRED") {
        throw new MailError("AUTH_REQUIRED", error.message + " For MFA or other manual steps, run npm run login -- --interactive.");
      }
      throw error;
    }
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
          this.clearSignInCooldown();
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

  private clearSignInCooldown() {
    this.signInBlockedUntil = 0;
    this.signInBlockedReason = undefined;
  }

  /** True when an operation failed because Microsoft signed the browser out and recovery has not been tried yet. */
  private async sessionLost(page: Page, error: unknown) {
    if (this.recoveryTried || page.isClosed()) return false;
    if (error instanceof SessionLostError) return true;
    // AUTH_REQUIRED means recovery already ran and failed; other errors may be a side effect of a redirect to sign-in.
    if (error instanceof MailError && error.code === "AUTH_REQUIRED") return false;
    return isMicrosoftLogin(page) || await this.signInPrompt(page);
  }

  /** Signs back in when the page is on a Microsoft sign-in page. Returns without effect on any other page. */
  private async recoverSession(page: Page) {
    if (!isMicrosoftLogin(page)) return;
    this.recoveryTried = true;
    const username = process.env.OUTLOOK_USERNAME;
    // A remembered account usually signs in again without a secret.
    const resumed = await resumeSignIn(page, username).catch(() => "needs_user" as const);
    if (resumed === "signed_in") {
      console.error("outlook-mcp: resumed the Microsoft session without a password.");
      return;
    }
    let source;
    try {
      source = passwordSource();
    } catch (error) {
      throw new MailError("AUTH_REQUIRED", (error as Error).message);
    }
    if (source.kind === "file") {
      const report = inspectPasswordFile(source.path);
      if (!report.valid) throw new MailError("AUTH_REQUIRED", report.message ?? "Password file misconfigured.");
    }
    if (!username || source.kind === "none") {
      throw new MailError("AUTH_REQUIRED", sessionExpired +
        " To sign back in automatically, set OUTLOOK_USERNAME and OUTLOOK_PASSWORD (or OUTLOOK_PASSWORD_FILE) in .env.");
    }
    if (Date.now() < this.signInBlockedUntil) {
      throw new MailError("AUTH_REQUIRED", (this.signInBlockedReason ?? "Automatic sign-in is cooling down after a recent attempt.") + " " + signInHelp);
    }
    try {
      await reauthenticate(page, username, source);
      this.clearSignInCooldown();
      console.error("outlook-mcp: signed back in automatically with the configured credentials.");
    } catch (error) {
      if (error instanceof PasswordFileError) {
        this.clearSignInCooldown();
        throw new MailError("AUTH_REQUIRED", error.message);
      }
      // Browser errors can echo the password that was filled, so only our own messages are forwarded.
      if (error instanceof SignInError && error.problem !== "stuck") {
        // Retrying a rejected password risks locking the account; wait for npm run login or a long cooldown.
        const rejected = error.problem === "password_rejected" || error.problem === "account_rejected" || error.problem === "account_locked";
        this.signInBlockedUntil = Date.now() + (rejected ? 6 * 3600_000 : 15 * 60_000);
        this.signInBlockedReason = error.message + " Automatic sign-in is paused.";
        throw new MailError("AUTH_REQUIRED", error.message);
      }
      // Transient failures (slow pages, unexpected prompts) retry soon.
      this.signInBlockedUntil = Date.now() + 2 * 60_000;
      this.signInBlockedReason = "Automatic sign-in is cooling down after a recent attempt.";
      throw new MailError("AUTH_REQUIRED", "Automatic sign-in failed. " + signInHelp);
    }
  }

  private async signInPrompt(page: Page) {
    return page.getByRole("link", { name: /^sign in$/i }).or(page.getByRole("button", { name: /^sign in$/i }))
      .or(page.getByText(/session (?:has )?expired|signed out|sign in again/i)).first().isVisible().catch(() => false);
  }

  /**
   * Marks a tab as belonging to another read-only surface so the next mail operation
   * does not reuse it. Calendar holds its own tab; the mailbox tab stays on mail.
   */
  holdPage(page: Page) {
    if (this.reserved.has(page)) return;
    this.reserved.add(page);
    page.once("close", () => this.reserved.delete(page));
  }

  /**
   * Runs read-only work in the signed-in browser after the mailbox session is ready.
   * Serialized with mail on the same queue, with the same auth recovery and error codes.
   */
  useBrowser<T>(work: (context: BrowserContext) => Promise<T>, signal?: AbortSignal) {
    return this.run(async page => {
      await this.ensureMailbox(page);
      if (!this.context) throw new MailError("BROWSER_ERROR", "The browser is not open.");
      this.checkpoint();
      return work(this.context);
    }, signal);
  }

  /**
   * Navigates to Outlook when needed, recovers sign-in and returns the search box.
   * Outlook often loads on its own origin and then redirects to Microsoft sign-in a moment later,
   * so this keeps watching for either the search box or a sign-in page until one settles.
   * `force` reloads Outlook first, used after a session ended mid-operation.
   */
  private async ensureMailbox(page: Page, force = false) {
    let blockedNavigation = false;
    const watch = (request: import("playwright").Request) => {
      const origin = safeOrigin(request.url());
      if (request.isNavigationRequest() && request.frame() === page.mainFrame()
          && !outlookOrigins.has(origin) && !signInOrigins.has(origin)) blockedNavigation = true;
    };
    const ignoreBlocked = (error: unknown) => {
      if (!blockedNavigation && !isMicrosoftLogin(page)) throw error;
    };
    const load = async (how: "goto" | "reload") => {
      if (how === "goto") await page.goto(mailbox.href, { waitUntil: "domcontentloaded", timeout: 30000 }).catch(ignoreBlocked);
      else await page.reload({ waitUntil: "domcontentloaded", timeout: 30000 }).catch(ignoreBlocked);
      this.loadedAt = Date.now();
      this.lastSearchQuery = undefined;
    };
    page.on("request", watch);
    try {
      await this.recoverSession(page);
      if (page.url() !== "about:blank" && interactiveMode() && !outlookOrigins.has(safeOrigin(page.url()))) {
        throw new MailError("AUTH_REQUIRED", loginWindowOpen + " Complete sign-in in that window.");
      }
      const stale = this.loadedAt > 0 && Date.now() - this.loadedAt > (this.options.reloadMs ?? 6 * 3600_000);
      if (force || stale || !outlookOrigins.has(safeOrigin(page.url()))) await load("goto");

      const modernSearch = page.locator("#topSearchInput");
      const search = modernSearch.or(page.getByRole("searchbox")).or(page.getByRole("textbox", { name: "Search", exact: true }))
        .or(page.locator('input[placeholder="Search"]')).first();
      const readyMs = this.options.readyMs ?? 20000;
      let reloaded = false;
      let prompted = false;
      let recoveries = 0;
      // Each pass waits for the search box or a sign-in page, then handles whichever appeared.
      for (;;) {
        this.checkpoint();
        if (isMicrosoftLogin(page)) {
          if (++recoveries > 2) throw new MailError("AUTH_REQUIRED", sessionExpired);
          await this.recoverSession(page);
          if (!outlookOrigins.has(safeOrigin(page.url()))) await load("goto");
        }
        const outcome = await waitForSurface(page, search, readyMs, () => this.checkpoint());
        if (outcome === "ready" && !isMicrosoftLogin(page)) break;
        if (outcome === "login" || isMicrosoftLogin(page)) continue;
        if (page.url().startsWith("chrome-error:") && !blockedNavigation) {
          throw new MailError("NETWORK_ERROR", "Could not reach Outlook. Check the network connection and try again.");
        }
        if (blockedNavigation) throw new MailError("AUTH_REQUIRED", sessionExpired);
        if (!prompted && await this.signInPrompt(page)) {
          // Outlook's own "session expired" prompt. Loading the mailbox again sends it to Microsoft sign-in.
          prompted = true;
          await load("goto");
          continue;
        }
        if (!reloaded && outlookOrigins.has(safeOrigin(page.url()))) {
          // Outlook occasionally shows an error, update or session page. One reload usually clears it.
          reloaded = true;
          await load("reload");
          continue;
        }
        if (await this.signInPrompt(page)) throw new MailError("AUTH_REQUIRED", sessionExpired);
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
    this.checkpoint();
    const existing = this.lastSearchQuery === query && await this.showingSearch(page, query)
      ? await this.extractRows(this.resultList(page))
      : [];
    if (existing.length) return existing;
    this.lastSearchQuery = undefined;
    const { search, modernSearch } = await this.ensureMailbox(page);
    const modern = await modernSearch.isVisible();
    if (modern) {
      await modernSearch.fill(query);
      if (await modernSearch.inputValue() !== query) {
        await modernSearch.click();
        await modernSearch.press("Control+a");
        await modernSearch.pressSequentially(query, { delay: 20 });
      }
      if (await modernSearch.inputValue() !== query) throw new MailError("UI_CHANGED", "Outlook did not accept the search query.");
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
    this.checkpoint();
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
    await this.scrollList(page, "top");
    add(await this.extractRows(this.resultList(page)));
    let stuck = 0;
    for (let step = 0; step < 50 && this.hitsFrom(collected).length < limit; step++) {
      const moved = await this.scrollList(page, "next");
      if (!moved) break;
      const before = collected.length;
      add(await this.extractRows(this.resultList(page)));
      if (collected.length === before) {
        if (++stuck >= 3) break;
      } else stuck = 0;
    }
    return collected;
  }

  search(query: string, limit: number, signal?: AbortSignal) {
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
    }, signal);
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
    this.checkpoint();
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
    this.checkpoint();
    const clicked = await candidate.evaluate(async (node, { s, fingerprint }) => {
      const text = (selector: string) => node.querySelector(selector)?.textContent?.trim() ?? "";
      const timeNode = node.querySelector("time");
      const time = timeNode?.getAttribute("datetime") ?? timeNode?.textContent?.trim()
        ?? node.querySelector(".qq2gS[title]")?.getAttribute("title") ?? "";
      let key: string | null = null;
      for (const name of s.keyAttributes) {
        const value = node.getAttribute(name) || node.querySelector("[" + name + "]")?.getAttribute(name);
        if (value) { key = value; break; }
      }
      if (JSON.stringify([key, text(s.sender), text(s.subject), time, text(s.preview)]) !== fingerprint) return false;
      const subject = node.querySelector<HTMLElement>(s.subject);
      if (!subject || !node.isConnected) return false;
      let scrollable = node.parentElement;
      while (scrollable && !(scrollable.scrollHeight > scrollable.clientHeight + 4
          && /auto|scroll/.test(getComputedStyle(scrollable).overflowY))) scrollable = scrollable.parentElement;
      const identity = () => scrollable ? JSON.stringify(Array.from(scrollable.querySelectorAll(s.results), row =>
        [row.getAttribute("data-item-id"), row.getAttribute("data-convid"), row.querySelector(s.subject)?.textContent])) : "";
      const before = scrollable?.scrollTop;
      const windowBefore = identity();
      subject.click();
      if (scrollable && scrollable.scrollTop !== before && identity() === windowBefore) {
        await new Promise<void>(resolve => {
          let frame = 0;
          const finish = () => { clearTimeout(timer); cancelAnimationFrame(frame); resolve(); };
          const check = () => { if (identity() !== windowBefore) finish(); else frame = requestAnimationFrame(check); };
          const timer = setTimeout(finish, 250);
          check();
        });
      }
      return true;
    }, { s: selectors, fingerprint: row.fingerprint });
    if (!clicked) throw new MailError("MESSAGE_CHANGED", "Search results changed before opening the message.");
    return { row, wasUnread };
  }

  // Moves the list's scroll container: "top" resets it, "next" advances most of a screen. Returns false at the end.
  private scrollList(page: Page, step: "top" | "next") {
    this.checkpoint();
    return searchRegion(page).evaluate(async (root, { step, results }) => {
      const scrollable = [root, ...root.querySelectorAll("*")].find(element =>
        element.scrollHeight > element.clientHeight + 4 && /auto|scroll/.test(getComputedStyle(element).overflowY));
      if (!scrollable) return false;
      const identity = () => JSON.stringify(Array.from(root.querySelectorAll(results)).map(node =>
        [node.getAttribute("data-item-id"), node.getAttribute("data-convid"), node.id, node.textContent]));
      const windowBefore = identity();
      const before = scrollable.scrollTop;
      scrollable.scrollTop = step === "top" ? 0 : before + scrollable.clientHeight * 0.8;
      if (scrollable.scrollTop === before) return false;
      await new Promise<void>(resolve => {
        let frame = 0;
        const finish = () => { clearTimeout(timer); cancelAnimationFrame(frame); resolve(); };
        const check = () => { if (identity() !== windowBefore) finish(); else frame = requestAnimationFrame(check); };
        const timer = setTimeout(finish, 250);
        check();
      });
      return true;
    }, { step, results: selectors.results }).catch(() => false);
  }

  /**
   * Outlook virtualizes the message list: only rows near the viewport exist in the page, and opening
   * messages scrolls the list. Look at the rendered rows, then page through the list from the top.
   */
  private async scrollToFind(page: Page, find: (rows: Row[]) => Row[]) {
    const visibleRows = () => this.resultList(page);
    let matches = find(await this.extractRows(visibleRows()));
    if (matches.length) return matches;
    await this.scrollList(page, "top");
    for (let steps = 0; steps < 25; steps++) {
      matches = find(await this.extractRows(visibleRows()));
      if (matches.length) return matches;
      const moved = await this.scrollList(page, "next");
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

  async read(id: string, maxBodyChars = 12000, signal?: AbortSignal) {
    const reference = this.reference(id);
    return this.run(page => this.readOnPage(page, reference, id, maxBodyChars), signal);
  }

  readMails(ids: string[], maxBodyChars = 12000, signal?: AbortSignal) {
    if (!ids.length || ids.length > 10) throw new MailError("INVALID_RANGE", "Pass 1 to 10 message IDs.");
    return this.run(async page => {
      const results = [];
      for (const message_id of ids) {
        this.checkpoint();
        try {
          results.push({ message_id, data: await this.readOnPage(page, this.reference(message_id), message_id, maxBodyChars) });
        } catch (error) {
          this.checkpoint();
          if (await this.sessionLost(page, error)) throw error;
          results.push({ message_id, error: errorInfo(this.classify(error)) });
        }
      }
      return { results };
    }, signal);
  }

  private async readOnPage(page: Page, reference: Reference, id: string, maxBodyChars: number) {
      if (!Number.isInteger(maxBodyChars) || maxBodyChars < 1 || maxBodyChars > 12000) {
        throw new MailError("INVALID_RANGE", "max_body_chars must be an integer from 1 to 12000.");
      }
      this.checkpoint();
      const opened = await this.openMessage(page, reference);
      try {
        this.checkpoint();
        return await this.extract(page, opened.row, id, maxBodyChars);
      } finally {
        if (opened.wasUnread) await this.restoreUnread(page);
      }
  }

  /**
   * Downloads through Outlook's own command; consume must finish before the temporary file is deleted.
   * `attachment` is a zero-based index into attachment_names or an exact file name.
   */
  async downloadAttachment<T extends Record<string, unknown>>(id: string, attachment: number | string, maxBytes: number,
    consume: (file: { path: string; name: string; size_bytes: number; head: Buffer }) => Promise<T>, signal?: AbortSignal) {
    const reference = this.reference(id);
    return this.run(async page => {
      const opened = await this.openMessage(page, reference);
      const row = opened.row;
      try {
        const pane = this.pane(page, row);
        await pane.locator(selectors.attachment).first().waitFor({ timeout: 8000 }).catch(() => {
          throw new MailError("NO_ATTACHMENTS", "This message has no downloadable attachments.");
        });
        const metadata = await pane.evaluate(mailMetadata, { s: selectors, row });
        const names = metadata.attachment_names;
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
        const visibleSize = metadata.attachments[index].size_bytes;
        if (visibleSize !== null && visibleSize > maxBytes) throw new MailError("ATTACHMENT_TOO_LARGE", "The visible attachment size exceeds the configured limit.");
        this.checkpoint();
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
          this.checkpoint();
          if (await download.failure()) throw new MailError("DOWNLOAD_FAILED", "Outlook could not download the attachment.");
          const path = await download.path();
          const { size } = await stat(path);
          if (size > maxBytes) {
            throw new MailError("ATTACHMENT_TOO_LARGE", "The attachment is " + Math.ceil(size / 1048576) + " MB, over the " +
              Math.floor(maxBytes / 1048576) + " MB limit.");
          }
          this.checkpoint();
          const name = names[index] || download.suggestedFilename();
          const fd = await open(path, "r");
          const head = Buffer.alloc(Math.min(size, 8192));
          try { await fd.read(head, 0, head.length, 0); } finally { await fd.close(); }
          const payload = await consume({ path, name, size_bytes: size, head });
          return { message_id: id, name, size_bytes: size, ...payload };
        } finally {
          await download.delete().catch(() => {});
        }
      } finally {
        if (opened.wasUnread) await this.restoreUnread(page);
      }
    }, signal);
  }

  private async extract(page: Page, row: Row, id: string, maxBodyChars: number) {
    this.checkpoint();
    const pane = this.pane(page, row);
    await pane.waitFor({ state: "visible" });
    const findBody = async () => {
      for (const selector of selectors.body) {
        const present = pane.locator(selector);
        if (await present.evaluateAll(nodes => nodes.some(node => {
          for (let current: Element | null = node; current; current = current.parentElement) {
            const style = getComputedStyle(current);
            if (current.hasAttribute("hidden") || current.getAttribute("aria-hidden") === "true"
                || style.display === "none" || style.visibility === "hidden") return false;
          }
          return true;
        }))) return present;
      }
    };
    let bodies: Locator | undefined;
    const started = Date.now();
    while (!bodies && Date.now() - started < 8000) {
      this.checkpoint();
      bodies = await findBody();
      if (!bodies) await page.waitForTimeout(150);
    }
    if (!bodies) throw new MailError("UI_CHANGED", "Could not find the message body in the reading pane.");
    // Outermost bodies only. A conversation shows one body per expanded message.
    const extracted = await bodies.evaluateAll((nodes, maxBodyChars) => {
      const eligible = nodes.filter(node => {
        for (let current: Element | null = node; current; current = current.parentElement) {
          const style = getComputedStyle(current);
          if (current.hasAttribute("hidden") || current.getAttribute("aria-hidden") === "true"
              || style.display === "none" || style.visibility === "hidden") return false;
        }
        return true;
      });
      const roots = eligible.filter(node => !eligible.some(other => other !== node && other.contains(node)));
      const parts: BodyPart[] = [];
      let committed = 0;
      let separator = false;
      let truncated = false;
      for (const [index, root] of roots.entries()) {
        if (truncated) break;
        if (index) separator = true;
        let buffer = "";
        let buffered = 0;
        let whitespace = "";
        const overhead = () => parts.length ? separator ? 7 : 2 : 0;
        const append = (text: string) => {
          const budget = Math.max(1, maxBodyChars + 1 - committed - overhead());
          for (const char of text) {
            if (char === " ") {
              if (buffer && !whitespace.includes("\n")) whitespace = " ";
              continue;
            }
            if (char === "\n") {
              if (buffer) whitespace = whitespace.includes("\n") ? "\n\n" : "\n";
              continue;
            }
            if (buffered >= budget) { truncated = true; break; }
            buffer += whitespace + char;
            buffered += whitespace.length + 1;
            whitespace = "";
            if (buffered > budget) { truncated = true; break; }
          }
        };
        const flush = (kind: BodyPart["kind"]) => {
          if (buffer) {
            committed += overhead();
            if (separator && parts.length) parts.push({ kind: "separator", text: "" });
            separator = false;
            parts.push({ kind, text: buffer });
            const clean = buffer.split("\n").map(line => line.replace(/[ \t ]+/g, " ").trim()).join("\n")
              .replace(/\n{3,}/g, "\n\n").trim();
            const safe = clean.replace(/[\\`\[\]<>]/g, "\\$&")
              .replace(/^(\s*)(#{1,6}|[-+*]|\d+[.)])(?=\s)/gm, "$1\\$2")
              .replace(/^(\s*)(={3,}|-{3,}|_{3,}|\*{3,})\s*$/gm, "$1\\$2");
            const rendered = kind === "bullet" ? "- " + safe.replace(/\n/g, "\n  ")
              : kind === "quote" ? safe.split("\n").map(line => "> " + line).join("\n") : safe;
            committed += Array.from(rendered).length;
            if (committed > maxBodyChars) truncated = true;
          }
          buffer = "";
          buffered = 0;
          whitespace = "";
        };
        const walk = (node: Node, kind: BodyPart["kind"]) => {
          for (const child of node.childNodes) {
            if (truncated) break;
            if (child.nodeType === Node.TEXT_NODE) {
              append((child.textContent ?? "").replace(/\s+/g, " "));
              continue;
            }
            if (child instanceof HTMLImageElement) {
              const alt = child.alt.replace(/\s+/g, " ").trim();
              if (alt.length > 1 && !child.hidden) append(" (image: " + alt + ") ");
              continue;
            }
            if (!(child instanceof HTMLElement) || child.hidden || child.getAttribute("aria-hidden") === "true" ||
                ["SCRIPT", "STYLE", "HEAD", "TITLE", "TEMPLATE", "svg"].includes(child.tagName)) continue;
            const style = getComputedStyle(child);
            if (style.display === "none" || style.visibility === "hidden") continue;
            if (child.tagName === "BR") {
              append("\n");
              continue;
            }
            // Data tables become "a | b" rows. Layout tables (one cell or nested tables) are walked normally.
            if (child instanceof HTMLTableRowElement && child.cells.length > 1 && !child.querySelector("table")) {
              flush(kind);
              for (const cell of child.cells) {
                const value = cell.innerText.replace(/\s+/g, " ").trim();
                if (value) append((buffer ? " | " : "") + value);
                if (truncated) break;
              }
              flush("row");
              continue;
            }
            const block = !style.display.startsWith("inline");
            const inner = child.tagName === "LI" ? "bullet" : child.tagName === "BLOCKQUOTE" ? "quote"
              : child.tagName === "HR" ? "separator" : kind;
            if (block) flush(kind);
            if (inner === "separator" && kind !== "separator") {
              separator = true;
            }
            else walk(child, inner);
            if (block) flush(inner);
          }
        };
        walk(root, "paragraph");
        flush("paragraph");
      }
      return { parts, truncated };
    }, maxBodyChars);
    // Attachment cards can render after the body. Wait for them only when the row says there are some.
    if (row.hasAttachments) await pane.locator(selectors.attachment).first().waitFor({ timeout: 8000 }).catch(() => {});
    this.checkpoint();
    const metadata = await pane.evaluate(mailMetadata, { s: selectors, row });
    const body = formatBody(extracted.parts, maxBodyChars);
    return { message_id: id, ...metadata, received_at: row.time || null,
      ...body, truncated: body.truncated || extracted.truncated };
  }

  async close() {
    await this.shutdownBrowser();
  }
}
