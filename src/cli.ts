import { displayVariables, profileDir } from "./browser.js";
import { callControl, controlSocketPath, ServerNotRunning } from "./control.js";
import { MailError, Outlook } from "./outlook.js";

const command = process.argv[2];
const display = Object.fromEntries(displayVariables.flatMap(key => process.env[key] ? [[key, process.env[key]!]] : []));

function fail(message: string): never {
  console.error(message);
  process.exit(1);
}

// Talks to the running server when there is one, so the profile is never opened twice.
async function viaServer<T>(method: "GET" | "POST", path: string, body?: unknown) {
  try {
    return await callControl<T>(method, path, body);
  } catch (error) {
    if (error instanceof ServerNotRunning) return undefined;
    fail("Server error: " + (error instanceof Error ? error.message : error));
  }
}

function describe(status: Record<string, unknown>) {
  const auth = { signed_in: "signed in", signed_out: "signed out (run npm run login)", unknown: "unknown" }[String(status.auth)];
  console.log("Mailbox:  " + auth);
  console.log("Browser:  " + status.browser + (status.busy ? " (" + status.busy + ")" : ""));
  if (status.last_success_at) console.log("Last OK:  " + status.last_success_at);
  const error = status.last_error as { code: string; message: string; at: string } | null;
  if (error) console.log("Last err: " + error.code + " at " + error.at + " — " + error.message);
}

async function login() {
  if (!display.DISPLAY && !display.WAYLAND_DISPLAY) fail("Run npm run login from a graphical session so the sign-in window can open.");
  console.log("Opening a sign-in window. Complete Microsoft sign-in and MFA, and choose \"Stay signed in\" if asked.");
  console.log("The window closes by itself once Outlook loads.");
  const remote = await viaServer<{ result: string }>("POST", "/login", { display });
  let result = remote?.result;
  if (!remote) {
    try {
      result = await new Outlook().signIn({ display });
    } catch (error) {
      fail(error instanceof MailError ? error.message : String(error));
    }
  }
  if (result === "signed_in") console.log("Signed in" + (remote ? ". The running server will use the new session." : "."));
  else fail(result === "cancelled" ? "Sign-in window closed before Outlook loaded." : "Sign-in timed out after 10 minutes.");
}

async function logout() {
  const remote = await viaServer<{ profile_deleted: string }>("POST", "/logout");
  if (remote) return console.log("Signed out. Deleted browser profile " + remote.profile_deleted + ".");
  try {
    const { profile_deleted } = await new Outlook().logout();
    console.log("Signed out. Deleted browser profile " + profile_deleted + ".");
  } catch (error) {
    fail(error instanceof MailError ? error.message : String(error));
  }
}

async function status() {
  const check = process.argv.includes("--check");
  const remote = await viaServer<Record<string, unknown>>(check ? "POST" : "GET", check ? "/check" : "/status");
  if (!remote) {
    console.log("Server:   not running (control socket " + controlSocketPath() + ")");
    console.log("Profile:  " + profileDir());
    return;
  }
  console.log("Server:   running");
  describe(remote);
}

const commands: Record<string, () => Promise<void>> = { login, logout, status };
if (!Object.hasOwn(commands, command)) fail("Usage: node dist/cli.js login | logout | status [--check]");
try {
  await commands[command]();
} catch (error) {
  fail(error instanceof MailError ? error.code + ": " + error.message : String(error));
}
