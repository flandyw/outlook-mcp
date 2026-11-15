import { closeSync, constants, fstatSync, lstatSync, openSync, readFileSync, readSync, readdirSync } from "node:fs";
import { basename, dirname, join } from "node:path";
import type { Page } from "playwright";
import { outlookOrigins, safeOrigin, signInOrigins } from "./browser.js";

export function isMicrosoftLogin(page: Page) {
  return signInOrigins.has(safeOrigin(page.url()));
}

// Microsoft's classic (idSIButton9) and newer (data-testid primaryButton) sign-in pages.
const submitSelector = [
  'input[type="submit"]#idSIButton9:visible', 'button[type="submit"]#idSIButton9:visible',
  'button[type="submit"][data-testid="primaryButton"]:visible',
].join(", ");
const submit = (page: Page) => page.locator(submitSelector).first();
const emailInput = (page: Page) => page.locator('input[name="loginfmt"]:visible, input[type="email"]:visible').first();
const passwordInput = (page: Page) => page.locator('input[name="passwd"]:visible, input[type="password"]:visible').first();
const staySignedIn = (page: Page) => page.getByText(/^Stay signed in\??$/).first();
const keepSignedInBox = (page: Page) => page.locator('#KmsiCheckboxField:visible, input[name="DontShowAgain"]:visible').first();
// "Let's keep your account secure" / "Action required" interstitials that can be postponed.
const postpone = (page: Page) => page.locator('#btnAskLater:visible, a#iShowSkip:visible')
  .or(page.getByRole("button", { name: /^(?:ask later|skip for now|not now|maybe later)$/i }))
  .or(page.getByRole("link", { name: /^(?:ask later|skip for now|not now|maybe later)$/i })).first();
const passwordRejected = (page: Page) => page.locator('#passwordError:visible, [id="i0118Error"]:visible')
  .or(page.getByText(/your account or password is incorrect|incorrect password|password is incorrect/i)).first();
const accountRejected = (page: Page) => page.locator('#usernameError:visible')
  .or(page.getByText(/(?:that|this) (?:microsoft )?account doesn.t exist|we couldn.t find an account/i)).first();
const accountLocked = (page: Page) => page.getByText(/your account (?:has been|is) (?:locked|blocked)|account is locked/i).first();
const secondFactor = (page: Page) => page.locator('#idDiv_SAOTCAS_Title:visible, #idDiv_SAOTCC_Title:visible, input[name="otc"]:visible, #idRichContext_DisplaySign:visible')
  .or(page.getByText(/^(?:approve sign.in request|verify your identity|enter code|more information required)$/i)).first();

async function shown(locator: ReturnType<Page["locator"]>) {
  return locator.isVisible().catch(() => false);
}

function accountTiles(page: Page, username?: string) {
  const tiles = page.locator('[data-test-id*="@"]:visible');
  return username ? tiles.and(page.locator(`[data-test-id="${username.replace(/["\\]/g, "")}" i]`)) : tiles;
}

async function acceptStaySignedIn(page: Page) {
  if (!await shown(staySignedIn(page)) || !await shown(submit(page))) return false;
  // "Don't show this again" keeps Microsoft from asking on every renewal.
  const box = keepSignedInBox(page);
  if (await shown(box) && !await box.isChecked().catch(() => true)) await box.check({ timeout: 1500 }).catch(() => {});
  await submit(page).click({ timeout: 1500 });
  return true;
}

/**
 * Finish a sign-in that needs no secret: choose the remembered account, accept "Stay signed in?"
 * and postpone optional security prompts.
 * Returns "signed_in" when Outlook loads, or "needs_password" / "needs_user" when Microsoft wants more.
 */
export async function resumeSignIn(page: Page, username?: string, timeout = 15000) {
  const deadline = Date.now() + timeout;
  const clicked = new Set<string>();
  while (Date.now() < deadline) {
    if (outlookOrigins.has(safeOrigin(page.url()))) return "signed_in" as const;
    if (!isMicrosoftLogin(page)) return "needs_user" as const;
    if (await shown(passwordInput(page))) return "needs_password" as const;
    if (await shown(secondFactor(page))) return "needs_user" as const;
    if (!clicked.has("kmsi") && await acceptStaySignedIn(page)) {
      clicked.add("kmsi");
    } else if (!clicked.has("later") && await shown(postpone(page))) {
      clicked.add("later");
      await postpone(page).click({ timeout: 1500 });
    } else if (!clicked.has("tile")) {
      // Account picker. Choose the configured account, or the only remembered one.
      const wanted = accountTiles(page, username);
      if (await wanted.count() === 1) {
        clicked.add("tile");
        await wanted.click({ timeout: 1500 });
      } else if (await shown(emailInput(page))) {
        // An empty username form needs the configured account, which only reauthenticate types.
        return username ? "needs_password" as const : "needs_user" as const;
      }
    }
    await page.waitForTimeout(300);
  }
  return outlookOrigins.has(safeOrigin(page.url())) ? "signed_in" as const : "needs_user" as const;
}

export type SignInProblem = "password_rejected" | "account_rejected" | "account_locked" | "second_factor" | "stuck";

/** Microsoft refused or paused an automatic sign-in. Messages never include the password. */
export class SignInError extends Error {
  constructor(readonly problem: SignInProblem, message: string) {
    super(message);
    this.name = "SignInError";
  }
}

export type PasswordFileProblem = "missing" | "not_a_file" | "unreadable" | "empty" | "loose_permissions" | "too_large";

/** What an operator can know about OUTLOOK_PASSWORD_FILE. Never includes file contents. */
export interface PasswordFileReport {
  configured: boolean;
  /** True only when a file is configured and every check passes. */
  valid: boolean;
  path: string | null;
  problem: PasswordFileProblem | null;
  /** Similarly named entries in the same directory when the file is missing. Never opened. */
  hints: string[];
  message: string | null;
}

export class PasswordFileError extends Error {
  constructor(message: string, readonly problem: PasswordFileProblem) {
    super(message);
    this.name = "PasswordFileError";
  }
}

const notConfigured: PasswordFileReport = {
  configured: false, valid: false, path: null, problem: null, hints: [], message: null,
};

function displayPath(filePath: string) {
  return filePath.replace(/[\u0000-\u001f\u007f]/g, "?");
}

function formatMode(mode: number) {
  return "0" + (mode & 0o777).toString(8).padStart(3, "0");
}

function describeProblem(filePath: string, problem: PasswordFileProblem, hints: string[], mode?: number) {
  const target = "OUTLOOK_PASSWORD_FILE " + displayPath(filePath);
  switch (problem) {
    case "missing": {
      const hint = hints.length
        ? " Similarly named files in the same directory (not used automatically): " + hints.map(displayPath).join(", ") + "."
        : "";
      return "Password file misconfigured: " + target + " is missing." + hint;
    }
    case "not_a_file":
      return "Password file misconfigured: " + target + " is not a regular file.";
    case "unreadable":
      return "Password file misconfigured: " + target + " is unreadable.";
    case "empty":
      return "Password file misconfigured: " + target + " is empty.";
    case "loose_permissions":
      return "Password file misconfigured: " + target + " has loose permissions" +
        (mode === undefined ? "" : " (mode " + formatMode(mode) + ")") +
        "; expected 0600 and not group or world accessible.";
    case "too_large":
      return "Password file misconfigured: " + target + " is too large.";
  }
}

function invalid(filePath: string, problem: PasswordFileProblem, mode?: number): PasswordFileReport {
  const hints = problem === "missing" ? similarFiles(filePath) : [];
  return {
    configured: true, valid: false, path: filePath, problem, hints,
    message: describeProblem(filePath, problem, hints, mode),
  };
}

// Similar names share a short edit distance or a long common prefix, so a renamed
// file in the same directory can be mentioned. Hints are names only and are never opened.
function similarFiles(filePath: string) {
  const directory = dirname(filePath);
  const wanted = basename(filePath);
  let names: string[];
  try {
    names = readdirSync(directory);
  } catch {
    return [];
  }
  const found: { path: string; distance: number }[] = [];
  for (const name of names) {
    if (name === wanted || name === "." || name === ".." || name.length > 255 || /[\u0000-\u001f\u007f]/.test(name)) continue;
    const distance = nameDistance(name, wanted);
    if (distance === undefined) continue;
    const candidate = join(directory, name);
    try {
      if (lstatSync(candidate).isDirectory()) continue;
    } catch {
      continue;
    }
    found.push({ path: candidate, distance });
  }
  found.sort((left, right) => left.distance - right.distance || left.path.localeCompare(right.path));
  return found.slice(0, 5).map(item => item.path);
}

function nameDistance(candidate: string, wanted: string) {
  const left = candidate.toLowerCase();
  const right = wanted.toLowerCase();
  if (left === right) return 0;
  const distance = editDistance(left, right);
  const longer = Math.max(left.length, right.length);
  const shorter = Math.min(left.length, right.length);
  const limit = longer <= 4 ? 1 : longer <= 8 ? 2 : Math.floor(longer * 0.35);
  if (distance <= limit) return distance;
  let prefix = 0;
  while (prefix < shorter && left[prefix] === right[prefix]) prefix++;
  if (prefix >= 4 && prefix / shorter >= 0.8 && prefix / longer >= 0.5) return distance;
  return undefined;
}

function editDistance(left: string, right: string) {
  if (left === right) return 0;
  if (!left.length) return right.length;
  if (!right.length) return left.length;
  let previous = Array.from({ length: right.length + 1 }, (_, index) => index);
  for (let i = 1; i <= left.length; i++) {
    const next = [i];
    for (let j = 1; j <= right.length; j++) {
      const cost = left.charCodeAt(i - 1) === right.charCodeAt(j - 1) ? 0 : 1;
      next[j] = Math.min(next[j - 1] + 1, previous[j] + 1, previous[j - 1] + cost);
    }
    previous = next;
  }
  return previous[right.length];
}

type Opened = { fd: number; size: number; mode: number };
type OpenProblem = { problem: PasswordFileProblem; mode?: number };

function rejectOpen(fd: number, problem: PasswordFileProblem, mode?: number): OpenProblem {
  closeSync(fd);
  return mode === undefined ? { problem } : { problem, mode };
}

/** Open the configured path only. O_NOFOLLOW refuses a symlink swapped in after lstat. */
function openPasswordFile(filePath: string): Opened | OpenProblem {
  let info;
  try {
    info = lstatSync(filePath);
  } catch (error) {
    const code = (error as NodeJS.ErrnoException).code;
    return { problem: code === "ENOENT" || code === "ENOTDIR" ? "missing" : "unreadable" };
  }
  if (!info.isFile()) return { problem: "not_a_file" };
  const listedMode = info.mode & 0o777;
  if ((info.mode & 0o077) !== 0) return { problem: "loose_permissions", mode: listedMode };
  if ((listedMode & 0o400) === 0) return { problem: "unreadable", mode: listedMode };

  let fd: number;
  try {
    fd = openSync(filePath, constants.O_RDONLY | constants.O_NOFOLLOW | constants.O_NONBLOCK);
  } catch (error) {
    const code = (error as NodeJS.ErrnoException).code;
    if (code === "ENOENT" || code === "ENOTDIR") return { problem: "missing" };
    if (code === "ELOOP") return { problem: "not_a_file" };
    return { problem: "unreadable" };
  }
  try {
    const opened = fstatSync(fd);
    const mode = opened.mode & 0o777;
    if (!opened.isFile()) return rejectOpen(fd, "not_a_file");
    if ((opened.mode & 0o077) !== 0) return rejectOpen(fd, "loose_permissions", mode);
    if ((mode & 0o400) === 0) return rejectOpen(fd, "unreadable", mode);
    if (opened.size > 4096) return rejectOpen(fd, "too_large");
    return { fd, size: opened.size, mode };
  } catch {
    return rejectOpen(fd, "unreadable");
  }
}

function bodyIsEmpty(fd: number, size: number) {
  if (size === 0) return true;
  // One trailing newline is not a password. Anything longer still has a character after it is stripped.
  if (size > 2) return false;
  const buffer = Buffer.alloc(size);
  const bytes = readSync(fd, buffer, 0, size, 0);
  return buffer.subarray(0, bytes).toString("utf8").replace(/\r?\n$/, "") === "";
}

/**
 * Check OUTLOOK_PASSWORD_FILE without using any other path.
 * A missing file may name similar entries in its directory; those are hints, not substitutes.
 * The report has no file contents. Contents are read only for a 1–2 byte file, to detect a blank line.
 */
export function inspectPasswordFile(configured = process.env.OUTLOOK_PASSWORD_FILE): PasswordFileReport {
  const filePath = configured?.trim() ?? "";
  if (!filePath) return notConfigured;
  const opened = openPasswordFile(filePath);
  if ("problem" in opened) return invalid(filePath, opened.problem, opened.mode);
  let empty = false;
  let unreadable = false;
  try {
    empty = bodyIsEmpty(opened.fd, opened.size);
  } catch {
    unreadable = true;
  } finally {
    closeSync(opened.fd);
  }
  if (unreadable) return invalid(filePath, "unreadable");
  if (empty) return invalid(filePath, "empty");
  return { configured: true, valid: true, path: filePath, problem: null, hints: [], message: null };
}

function readPassword(filePath: string) {
  const report = inspectPasswordFile(filePath);
  if (!report.configured || !report.valid || !report.path) {
    throw new PasswordFileError(report.message ?? "Password file misconfigured.", report.problem ?? "missing");
  }
  const opened = openPasswordFile(report.path);
  if ("problem" in opened) throw new PasswordFileError(describeProblem(report.path, opened.problem, opened.problem === "missing" ? similarFiles(report.path) : [], opened.mode), opened.problem);
  try {
    const password = readFileSync(opened.fd, "utf8").replace(/\r?\n$/, "");
    if (!password) throw new PasswordFileError(describeProblem(report.path, "empty", []), "empty");
    return password;
  } catch (error) {
    if (error instanceof PasswordFileError) throw error;
    throw new PasswordFileError(describeProblem(report.path, "unreadable", []), "unreadable");
  } finally {
    closeSync(opened.fd);
  }
}

export type PasswordSource = { kind: "env" } | { kind: "file"; path: string } | { kind: "none" };

/** OUTLOOK_PASSWORD (in .env) or OUTLOOK_PASSWORD_FILE. Setting both is a configuration error. */
export function passwordSource(): PasswordSource {
  const env = process.env.OUTLOOK_PASSWORD;
  const file = process.env.OUTLOOK_PASSWORD_FILE?.trim();
  if (env && file) throw new PasswordFileError("Set only one of OUTLOOK_PASSWORD and OUTLOOK_PASSWORD_FILE.", "missing");
  if (env) return { kind: "env" };
  if (file) return { kind: "file", path: file };
  return { kind: "none" };
}

/** Summary for status output. Never includes the password. */
export function credentialsSummary() {
  let source: PasswordSource["kind"] | "conflict";
  try {
    source = passwordSource().kind;
  } catch {
    source = "conflict";
  }
  return { username: process.env.OUTLOOK_USERNAME || null, password_source: source };
}

function resolvePassword(source: string | { value: string } | PasswordSource) {
  if (typeof source === "string") return readPassword(source);
  if ("value" in source) return source.value;
  if (source.kind === "file") return readPassword(source.path);
  if (source.kind === "env" && process.env.OUTLOOK_PASSWORD) return process.env.OUTLOOK_PASSWORD;
  throw new PasswordFileError("No password configured. Set OUTLOOK_PASSWORD or OUTLOOK_PASSWORD_FILE.", "missing");
}

/**
 * Type the configured account and password into Microsoft's sign-in pages, then accept "Stay signed in?".
 * `source` is a password file path, a literal password, or a configured source (default: OUTLOOK_PASSWORD / OUTLOOK_PASSWORD_FILE).
 * Throws SignInError when Microsoft rejects the password or needs MFA, so the caller stops retrying.
 */
export async function reauthenticate(page: Page, username: string, source: string | { value: string } | PasswordSource = passwordSource(), timeout = 45000) {
  const deadline = Date.now() + timeout;
  // Read the configured file only. A similarly named file is never a fallback.
  const password = resolvePassword(source);

  let sentUsername = false;
  let sentPassword = false;
  const clicked = new Set<string>();
  while (Date.now() < deadline) {
    if (outlookOrigins.has(safeOrigin(page.url()))) return;
    if (!isMicrosoftLogin(page)) break;

    if (sentPassword && await shown(passwordRejected(page))) {
      throw new SignInError("password_rejected", "Microsoft rejected the configured password. Update OUTLOOK_PASSWORD or the password file, then run npm run login.");
    }
    if (sentUsername && await shown(accountRejected(page))) {
      throw new SignInError("account_rejected", "Microsoft did not recognise OUTLOOK_USERNAME. Check it, then run npm run login.");
    }
    if (await shown(accountLocked(page))) {
      throw new SignInError("account_locked", "Microsoft reports the account is locked. Unlock it, then run npm run login -- --interactive.");
    }
    if (await shown(secondFactor(page))) {
      throw new SignInError("second_factor", "Microsoft asked for MFA or another verification step. Run npm run login -- --interactive to finish it.");
    }

    const email = emailInput(page);
    const secret = passwordInput(page);
    const button = submit(page);
    const tile = accountTiles(page, username);
    if (!clicked.has("tile") && !sentPassword && await tile.count() === 1) {
      clicked.add("tile");
      await tile.click({ timeout: 1500 });
    } else if (!sentUsername && !sentPassword && await shown(email) && await shown(button)) {
      sentUsername = true;
      await email.fill(username, { timeout: 1500 });
      if (!isMicrosoftLogin(page)) break;
      await button.click({ timeout: 1500 });
    } else if (!sentPassword && await shown(secret) && await shown(button)
        && (sentUsername || clicked.has("tile") || await shown(page.getByText(username, { exact: true }).first()))) {
      sentPassword = true;
      await secret.fill(password, { timeout: 1500 });
      if (!isMicrosoftLogin(page)) break;
      await button.click({ timeout: 1500 });
    } else if (!clicked.has("kmsi") && await acceptStaySignedIn(page)) {
      clicked.add("kmsi");
    } else if (!clicked.has("later") && await shown(postpone(page))) {
      clicked.add("later");
      await postpone(page).click({ timeout: 1500 });
    }
    await page.waitForTimeout(300);
  }
  throw new SignInError("stuck", "Microsoft sign-in needs attention.");
}
