import { chromium } from "playwright";
import { constants } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import { access, mkdir, chmod } from "node:fs/promises";

export const mailbox = new URL(process.env.OUTLOOK_URL ?? "https://outlook.office.com/mail/");
export const outlookOrigins = new Set(["https://outlook.office.com", "https://outlook.office365.com", "https://outlook.cloud.microsoft", "https://outlook.live.com"]);
// Microsoft sign-in hosts. Add a federated identity provider (for example ADFS) with OUTLOOK_SIGNIN_ORIGINS.
export const signInOrigins = new Set([
  "https://login.microsoftonline.com", "https://login.microsoft.com", "https://login.live.com",
  "https://account.live.com", "https://autologon.microsoftazuread-sso.com", "https://device.login.microsoftonline.com",
  ...(process.env.OUTLOOK_SIGNIN_ORIGINS ?? "").split(",").map(origin => origin.trim()).filter(Boolean),
]);
if (!outlookOrigins.has(mailbox.origin)
    || !mailbox.pathname.startsWith("/mail") || mailbox.username || mailbox.password) {
  throw new Error("OUTLOOK_URL must point to an Outlook Web mailbox.");
}

export function safeOrigin(url: string) {
  try {
    return new URL(url).origin;
  } catch {
    return "null";
  }
}

export const profileDir = () => process.env.OUTLOOK_PROFILE ?? join(homedir(), ".local/state/outlook-mcp/chromium");

export class ProfileLockedError extends Error {
  constructor() {
    super("Another browser is using the Outlook profile. Close any npm run login window or stop the other server.");
    this.name = "ProfileLockedError";
  }
}

// Variables a visible sign-in window needs from the desktop session that asked for it.
export const displayVariables = ["DISPLAY", "WAYLAND_DISPLAY", "XAUTHORITY", "XDG_RUNTIME_DIR", "DBUS_SESSION_BUS_ADDRESS"] as const;

export async function launch(headless: boolean, display: Partial<Record<string, string>> = {}) {
  const profile = profileDir();
  await mkdir(profile, { recursive: true, mode: 0o700 });
  await chmod(profile, 0o700);
  let executablePath = process.env.CHROMIUM_PATH;
  if (!executablePath && process.platform === "linux") {
    try {
      await access(chromium.executablePath(), constants.X_OK);
    } catch {
      try {
        await access("/usr/bin/chromium", constants.X_OK);
        executablePath = "/usr/bin/chromium";
      } catch {
        // Let Playwright report how to install its bundled browser.
      }
    }
  }
  const env: Record<string, string> = {};
  for (const [key, value] of Object.entries({ ...process.env, ...display })) if (value !== undefined) env[key] = value;
  const context = await chromium.launchPersistentContext(profile, {
    headless,
    executablePath,
    env,
    // Downloads happen only when read_attachment clicks Download; files are deleted after reading.
    acceptDownloads: true,
    serviceWorkers: "block",
  }).catch(error => {
    if (/ProcessSingleton|SingletonLock|profile.*(?:in use|locked)|user data directory is already in use|Opening in existing browser session/i.test(String(error))) {
      throw new ProfileLockedError();
    }
    throw error;
  });
  if (headless || process.env.OUTLOOK_INTERACTIVE === "1") {
    await context.route("**/*", async route => {
      const request = route.request();
      const origin = safeOrigin(request.url());
      let frameOrigin = "null";
      try {
        frameOrigin = safeOrigin(request.frame().url());
      } catch {
        // Requests without a frame (for example from workers) keep the default.
      }
      if ((request.resourceType() === "image" && (headless || outlookOrigins.has(frameOrigin))) ||
          (headless && request.isNavigationRequest() && !outlookOrigins.has(origin) && !signInOrigins.has(origin))) {
        await route.abort();
        return;
      }
      await route.continue();
    });
  }
  return context;
}
