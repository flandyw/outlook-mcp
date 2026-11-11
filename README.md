# Outlook Web MCP

A read-only TypeScript MCP server with three tools backed by a dedicated Playwright Chromium profile. Search returns message summaries and temporary IDs. Read returns available message metadata, Markdown and attachment names. Attachments are returned as the original file for the model to read. Search covers the results currently rendered by Outlook, including only rows the adapter can identify.

This is an unofficial browser adapter, with no Microsoft Graph API registration
required. It is not affiliated with Microsoft. Opening messages may mark them as
read; “read-only” describes the exposed tools, not a guarantee of zero mailbox
side effects.

## Requirements

- Node.js **22.9 or newer** and npm.
- Chromium (Playwright-managed, or a local installation via `CHROMIUM_PATH`).
- A graphical session for interactive Microsoft sign-in and an Outlook Web account.
- Linux is the primary target; the service installer requires systemd. Control
  commands use Unix sockets, so native Windows is not supported. macOS is unverified.

## Tools

| Tool | Input | Returns |
| --- | --- | --- |
| `search_mail` | `query` (Outlook search syntax, e.g. `from:alice`, `subject:invoice`, `hasattachments:yes`), `limit` 1–25 | `message_id`, sender, subject, `received_at`, `has_attachments`, snippet for each rendered result |
| `read_mail` | `message_id` from `search_mail` | sender, `to`, `cc`, `received_at`, `attachment_names`, `body_markdown`, `truncated` |
| `read_attachment` | `message_id`, `attachment` (index into `attachment_names` or exact file name) | file `name`, `size_bytes`, `mime`, `kind`, plus the file itself: an image block for PNG, JPEG, GIF and WebP; an embedded text resource for UTF-8 text files (txt, csv, html, ics, json…); an embedded base64 resource for everything else (PDF, Word, Excel, PowerPoint, audio…) |

Attachments are fetched with Outlook's own **Download** command, read into memory, deleted straight away and returned unchanged; the server does not convert or extract anything, so it is up to your MCP client and model to read PDFs or Office files. Files over 25 MB are refused (`OUTLOOK_ATTACHMENT_MAX_MB`); large files make large responses, so lower the limit if your client struggles. Set `OUTLOOK_ATTACHMENTS=0` to remove the tool. Cloud (OneDrive/SharePoint) links cannot be downloaded.

Outlook renders roughly the first 10–20 results, so narrow the query rather than raising `limit`. When a conversation shows several expanded messages, `read_mail` returns each body separated by `---`. Images are blocked; their alt text appears as `(image: …)`. Bodies are capped at 12,000 characters.

## Local setup

Run these from the project directory on a machine with a graphical session for the first sign-in.

```bash
npm ci
npm run build
npx playwright install chromium   # optional on Linux: falls back to /usr/bin/chromium
umask 077
printf 'MCP_TOKEN=%s\n' "$(openssl rand -hex 32)" > .env
npm run login
```

`npm run login` opens a sign-in window. Complete Microsoft sign-in and MFA and choose **Stay signed in** when asked: that persistent sign-in is what lets the server run headless afterwards. The window closes by itself once Outlook loads. For Outlook.com, put `OUTLOOK_URL=https://outlook.live.com/mail/` in `.env` (the default is `https://outlook.office.com/mail/`).

In Outlook's mail settings, disable automatic marking as read where available, and verify with an unread test message before using the server on real mail. Use a dedicated profile: logout deletes its directory. The profile lives in `~/.local/state/outlook-mcp/chromium` (`OUTLOOK_PROFILE`). Keep it private.

## Run as a service (recommended)

```bash
npm run service:install
```

This installs and starts a systemd user service, `outlook-mcp`, which runs the server headless, restarts it on failure and starts it at login. Run `loginctl enable-linger` if it should keep running while you are logged out. Logs: `journalctl --user -u outlook-mcp -f`. Remove it with `npm run service:uninstall`. After `npm run build`, restart it with `systemctl --user restart outlook-mcp`.

The headless browser starts on the first request and closes after 20 idle minutes (`OUTLOOK_IDLE_MINUTES`, `0` keeps it open), so an idle server uses almost no memory. A long-open Outlook page is reloaded every 6 hours.

To run in a terminal instead: `npm start`.

The server binds to `127.0.0.1:8765` by default. Use `MCP_PORT` to change the port. Configure the MCP client for Streamable HTTP and pass `Authorization: Bearer <MCP_TOKEN>` using its secret settings. The local endpoint is `http://127.0.0.1:8765/mcp`. Clients that send an Origin header need that exact origin in `MCP_ALLOWED_ORIGINS`, a comma-separated list.

## Sign-in, sign-out and status

These commands talk to the running server through a private Unix socket (`$XDG_RUNTIME_DIR/outlook-mcp/control.sock`, mode 0600). When no server is running, they use the profile directly.

| Command | What it does |
| --- | --- |
| `npm run login` | Opens a visible sign-in window on your desktop, even when the server runs as a service. Mail requests answer `AUTH_REQUIRED` until it finishes. Times out after 10 minutes. |
| `npm run logout` | Closes the browser and deletes the browser profile, removing every saved Microsoft sign-in. |
| `npm run status` | Shows sign-in state, whether the browser is running, and the last success or error. Add `-- --check` to load Outlook and test the session now. |

When Microsoft ends the session, the server first tries to resume it without a secret: it picks the remembered account in the account picker and accepts "Stay signed in?". If Microsoft still wants a password or MFA, mail requests return `AUTH_REQUIRED`, the journal logs it, and a desktop notification asks you to run `npm run login` (`OUTLOOK_NOTIFY=0` turns notifications off). Login and logout are only available on the local socket, never on the MCP port, so a tailnet client cannot trigger them.

Federated sign-in pages outside Microsoft (for example an organisation's ADFS) are blocked in headless mode. Add their origins to `OUTLOOK_SIGNIN_ORIGINS`, comma-separated.

## Keep a visible browser open instead

If Microsoft will not keep you signed in between browser restarts, set `OUTLOOK_INTERACTIVE=1` in `.env` and run `npm start` in a terminal (the service forces headless mode). The server opens a window; complete sign-in there and leave it open while using the MCP. If you close it or restart the server, sign in again.

## Automatically sign back in

For accounts that use a password without MFA, add these two settings to `.env` and restart the server.

```dotenv
OUTLOOK_USERNAME=your-microsoft-sign-in-address
OUTLOOK_PASSWORD_FILE=/absolute/path/to/a/private/outlook-password
```

Create the password file outside the project with permissions `0600` and put only the Microsoft password on its first line. This is a plaintext secret readable by your OS account. Use a private local location and exclude it from syncing and backups.

When a mail request reaches a Microsoft password form that the account picker cannot skip, the worker tries signing in once and then resumes the request. If sign-in fails or Microsoft requires another step, the tool returns `AUTH_REQUIRED` and waits five minutes before retrying automatically; `npm run login` works at any time.

## Access from another tailnet device

Restrict access to the host's HTTPS port in your tailnet policy before publishing the service. Run this on the browser host while the server is running.

```bash
tailscale serve --bg http://127.0.0.1:8765
tailscale serve status
```

Use the reported HTTPS address with `/mcp` appended and the same bearer header. If HTTPS certificates are not enabled for the tailnet, serve plain HTTP on the same port instead (traffic is still encrypted by WireGuard): `tailscale serve --bg --http=8765 http://127.0.0.1:8765`, then use `http://<machine>.<tailnet>.ts.net:8765/mcp`. Address the server by its MagicDNS name; Serve does not answer on the bare tailnet IP. Remove it with `tailscale serve --http=8765 off`. Tailscale Serve requires a client with access to the tailnet. Serve proxies the service on that address, so check existing Serve routes before applying the command. This project does not change your Tailscale policy or install a background service.

## Verification

```bash
npm test
```

Tests use Playwright's Chromium, falling back to `/usr/bin/chromium` on Linux (override with `CHROMIUM_PATH`). The tests exercise Chromium fixtures (including expired sessions, the account picker, crashes, hangs, network loss and sign-in/sign-out), the authenticated MCP handshake and the control socket. The fixture does not validate selectors against your signed-in Outlook layout. After sign-in, search for a known unread email, call `read_mail` with its returned ID, and verify its subject, body and unread state in Outlook.

A search result ID expires after 30 minutes or on server restart. Error codes:

- `UNKNOWN_MESSAGE` / `MESSAGE_CHANGED`: the ID expired or the result moved. Search again.
- `AUTH_REQUIRED`: the browser needs a fresh sign-in (`npm run login`), or a sign-in window is open.
- `SIGNED_OUT`: `npm run logout` is running.
- `NETWORK_ERROR`: Outlook could not be reached.
- `TIMEOUT`: Outlook was slow. If one request runs longer than 90 seconds, the browser is reset. Retry.
- `BROWSER_RESTARTED`: the browser crashed or its window closed. The next request starts a new one.
- `PROFILE_LOCKED`: another browser (an old `npm start`, or a direct `npm run login`) is using the profile.
- `UI_CHANGED`: the adapter could not identify a search region, row or reading pane. Outlook error pages get one automatic reload first.
- `BROWSER_ERROR`: anything else.

For `TIMEOUT`, `UI_CHANGED` and `BROWSER_ERROR`, the server prints the failing Playwright step to stderr. Outlook's DOM differs between accounts and changes over time. The selectors are in `src/outlook.ts`, and the fixture in `test/mail.test.mjs` mirrors the layout they were checked against (`outlook.cloud.microsoft`, November 2026).

Email fields are untrusted source data. The worker blocks images during mailbox operations and does not follow email links. It downloads an attachment only when `read_attachment` asks for it, never opens or runs it, and deletes it after returning it. Attachment content is untrusted in the same way as email bodies. The server does not store email bodies. Your MCP client and model provider may receive and retain returned content according to their own settings.

## Development and security

See [CONTRIBUTING.md](CONTRIBUTING.md) for development checks and sanitized bug
reports, [SECURITY.md](SECURITY.md) for the security model, and
[.env.example](.env.example) for configuration settings. GitHub CI runs the fixture
suite on Node.js 22, 24 and 26. These checks do not validate a live mailbox.

## License

[ISC](LICENSE).
