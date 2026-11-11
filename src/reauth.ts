import { open } from "node:fs/promises";
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

export async function reauthenticate(page: Page, username: string, passwordFile: string) {
  const deadline = Date.now() + 20000;
  const file = await open(passwordFile, "r");
  let password: string;
  try {
    const stat = await file.stat();
    if (!stat.isFile() || (stat.mode & 0o077) !== 0 || stat.size > 4096) {
      throw new Error("Password file must be a private regular file.");
    }
    password = (await file.readFile("utf8")).replace(/\r?\n$/, "");
    if (!password) throw new Error("Password file is empty.");
  } finally {
    await file.close();
  }

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
