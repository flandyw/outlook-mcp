import { chromium } from "playwright";

// Same fallback as the server: Playwright's browser, else system Chromium on Linux.
export async function launchChromium() {
  if (process.env.CHROMIUM_PATH) return chromium.launch({ executablePath: process.env.CHROMIUM_PATH });
  try {
    return await chromium.launch();
  } catch (error) {
    if (process.platform !== "linux") throw error;
    return chromium.launch({ executablePath: "/usr/bin/chromium" });
  }
}
