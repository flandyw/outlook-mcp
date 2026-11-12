import { closeSync, constants, fstatSync, lstatSync, openSync, readFileSync, readSync, readdirSync } from "node:fs";
import { basename, dirname, join } from "node:path";
import type { Page } from "playwright";
import { outlookOrigins, safeOrigin, signInOrigins } from "./browser.js";

export function isMicrosoftLogin(page: Page) {
  return signInOrigins.has(safeOrigin(page.url()));
}

const staySignedIn = (page: Page) => page.getByText(/^Stay signed in\?$/).first();
const submit = (page: Page) => page.locator('input[type="submit"]#idSIButton9:visible, button[type="submit"]#idSIButton9:visible').first();

/**
 * Finish a sign-in that needs no secret: choose the remembered account and accept "Stay signed in?".
 * Returns "signed_in" when Outlook loads, or "needs_password" / "needs_user" when Microsoft wants more.
 */
export async function resumeSignIn(page: Page, username?: string, timeout = 15000) {
  const deadline = Date.now() + timeout;
  const clicked = new Set<string>();
  while (Date.now() < deadline) {
    if (outlookOrigins.has(safeOrigin(page.url()))) return "signed_in" as const;
    if (!isMicrosoftLogin(page)) return "needs_user" as const;
    if (await page.locator('input[name="passwd"]:visible').count()) return "needs_password" as const;
    if (!clicked.has("kmsi") && await staySignedIn(page).isVisible() && await submit(page).isVisible()) {
      clicked.add("kmsi");
      await submit(page).click({ timeout: 1500 });
    } else if (!clicked.has("tile")) {
      // Account picker. Choose the configured account, or the only remembered one.
      const tiles = page.locator('[data-test-id*="@"]:visible');
      const wanted = username ? tiles.and(page.locator(`[data-test-id="${username.replace(/["\\]/g, "")}" i]`)) : tiles;
      if (await wanted.count() === 1) {
        clicked.add("tile");
        await wanted.click({ timeout: 1500 });
      } else if (await page.locator('input[name="loginfmt"]:visible').count() && !username) {
        return "needs_user" as const;
      }
    }
    await page.waitForTimeout(300);
  }
  return outlookOrigins.has(safeOrigin(page.url())) ? "signed_in" as const : "needs_user" as const;
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

export async function reauthenticate(page: Page, username: string, passwordFile: string) {
  const deadline = Date.now() + 20000;
  // Read the configured file only. A similarly named file is never a fallback.
  const password = readPassword(passwordFile);

  let sentUsername = false;
  let sentPassword = false;
  let acceptedPersistence = false;
  while (Date.now() < deadline) {
    if (outlookOrigins.has(safeOrigin(page.url()))) return;
    if (!isMicrosoftLogin(page)) break;

    const email = page.locator('input[name="loginfmt"]:visible');
    const secret = page.locator('input[name="passwd"]:visible');
    const submit = page.locator('input[type="submit"]#idSIButton9:visible, button[type="submit"]#idSIButton9:visible');
    if (!sentUsername && !sentPassword && await email.isVisible() && await submit.isVisible()) {
      sentUsername = true;
      await email.fill(username, { timeout: 1500 });
      if (!isMicrosoftLogin(page)) break;
      await submit.click({ timeout: 1500 });
    } else if (!sentPassword && await secret.isVisible() && await submit.isVisible()
        && (sentUsername || await page.getByText(username, { exact: true }).first().isVisible())) {
      sentPassword = true;
      await secret.fill(password, { timeout: 1500 });
      if (!isMicrosoftLogin(page)) break;
      await submit.click({ timeout: 1500 });
    } else if (sentPassword && !acceptedPersistence && await page.getByText("Stay signed in?", { exact: true }).isVisible()
        && await submit.isVisible()) {
      acceptedPersistence = true;
      await submit.click({ timeout: 1500 });
    }
    await page.waitForTimeout(300);
  }
  throw new Error("Microsoft sign-in needs attention.");
}
