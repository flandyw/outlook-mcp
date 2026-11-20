# Outlook Web MCP

A read-only TypeScript MCP server with eight tools backed by a dedicated Playwright Chromium profile. Search returns message summaries and message IDs. Read returns available message metadata, Markdown and indexed attachment metadata. Attachments are returned as temporary download links to the original file. Search scrolls the on-screen result list and returns only rows the adapter can identify. `mail_status` reports last-observed sign-in diagnostics and does not read mail. Calendar search reads events rendered in day view; read_event returns their details and notes.

This is an unofficial browser adapter, with no Microsoft Graph API registration
required. It is not affiliated with Microsoft. Opening a message can mark it as
read. When the row’s accessible name starts with “Unread” (or it has a badge
labeled exactly “Unread”) and Outlook shows a single **Mark as unread** button
outside the message body, the adapter clicks that button after a successful
read or attachment download. Buttons inside the message are ignored. If that
control is missing, the message may stay read. “Read-only” describes the
exposed tools: there is no send, reply, delete, move, or general flag action.

## Requirements

- Node.js **22.9 or newer** and npm.
- Chromium (Playwright-managed, or a local installation via `CHROMIUM_PATH`).
- A graphical session for interactive Microsoft sign-in and an Outlook Web account.
- Linux is the primary target; the service installer requires systemd. Control
  commands use Unix sockets, so native Windows is not supported. macOS is unverified.

## Tools

Mail/calendar queue jobs get one automatic retry for transient timeout, UI-change, network or browser failures, after reloading the mailbox and closing calendar tabs so the view is re-established. The original enqueue-time deadline and cancellation still apply. Auth/validation failures and storage pressure are not transient-view retries. Existing session-expiry recovery remains separate from auth-error reporting, and shares the single retry budget. Jobs stop being replayable before opening unread mail, starting a download, discarding a calendar editor, or after a successful batch item; unread restoration therefore runs once, not twice for a replay. Failures after those boundaries are reported without replay, including attachment-save failures.

`list_mail` bypasses the search index, which can lag fresh arrivals. It navigates to a live folder URL, clearing active search, and scans up to 50 scroll windows. Parseable dates are sorted newest first within their original dated slots; rows with unparseable dates keep their original positions, and equal dates retain list order. `scan_limit_reached` warns when scanning hit a bound or stopped making progress; `truncated` also covers omitted rows beyond `limit`. Only visible unread prefixes/badges are used for `unread_only`. Focused/Other and other Outlook view settings still determine which rows are exposed. Folder IDs work with `read_mail`, `read_mails`, and `read_attachment`, including after restart when they carry an Outlook key; they reopen the folder, never fall back to search. Drafts are list-only: reads return `DRAFT_UNAVAILABLE` rather than opening an editable composer. Search and folder rows include `unread`.

| Tool | Input | Returns |
| --- | --- | --- |
| `list_mail` | `folder` (default `inbox`; also `sentitems`, `drafts`, `archive`, `junkemail`, `deleteditems`), `limit` 1–40 (default 10), `unread_only` (default false) | search-shaped rows including `unread`, plus folder, scan/truncation flags; newest parseable dates first |
| `search_mail` | `query` (Outlook search syntax, e.g. `from:alice`, `subject:invoice`, `hasattachments:yes`), `limit` 1–40 | `message_id`, sender, subject, `received_at`, `has_attachments`, snippet for each identifiable result |
| `read_mail` | `message_id` from `search_mail`; optional `max_body_chars` 1–12,000 (default 12,000) | sender, `to`, `cc`, `received_at`, `attachment_names`, `attachments` (`index`, `name`, visible `size_bytes`/`type` or null), `body_markdown`, `truncated` |
| `read_mails` | `message_ids` (1–10 IDs); optional `max_body_chars` as above | `results` in input order, each with `message_id` and either `data` (read_mail result) or `error` |
| `read_attachment` | `message_id`, `attachment` (zero-based attachment index or exact full filename, case-insensitive; duplicate names require an index) | `name`, `size_bytes`, `mime`, `kind: download`, `download_url`, `expires_at`, and a `resource_link` content item |
| `mail_status` | none | `auth`, `browser`, `busy`, `last_success_at`, `last_error`, the automatic sign-in username and password source (`env`, `file` or `none`; never the password), and whether `OUTLOOK_PASSWORD_FILE` is configured and valid. `password_file_path` is included when configured. It does not read mail, sign in, or sign out. |
| `search_events` | `start` and `end` (`YYYY-MM-DD`, end inclusive, at most 14 days), optional `query`, `limit` 1–25 | `event_id`, title, `start`, `end`, `all_day`, location and `date` for each rendered event |
| `read_event` | `event_id` from `search_events` | times, location, attendees, recurrence, `body_markdown`, `truncated` |
| `read_events` | `event_ids` (1–10 IDs) | `results` in input order, each with `event_id` and either `data` (read_event result) or `error` |

Every tool declares an output schema and returns machine-readable `structuredContent` plus a human-readable text fallback (not a duplicate JSON dump). Errors retain `isError` and include `code`, `message`, and `retryable`. Batch item errors use the same fields. All tools are read-only and non-destructive; only local `mail_status` is advertised as idempotent. Mail reads can transiently mark mail as read; the existing best-effort unread restoration remains in place.

Attachments are fetched with Outlook's own **Download** command and moved into private temporary storage (disk-to-disk copy only if crossing filesystems), without buffering the whole file in Node. Only a small header is read to identify MIME type. The server does not convert, extract, open or run files. Files over 25 MiB are refused (`OUTLOOK_ATTACHMENT_MAX_MB`), using the visible size before download when available and checking the actual file size afterwards. Set `OUTLOOK_ATTACHMENTS=0` to remove the tool. Cloud (OneDrive/SharePoint) links cannot be downloaded.

Mail and calendar share one serialized browser queue. Each request's three-minute deadline begins at enqueue time and includes queue wait and browser launch. Cancelled/disconnected requests are skipped before they start, or stop at checkpoints between active steps, retaining unread/editor cleanup. MCP cancellation notifications match unique active request IDs across stateless POSTs; clients should use globally unique request IDs when sharing this endpoint. Ambiguous cancellation IDs are ignored rather than cancelling another client's work. Batches occupy one queue job; calendar batch reads are grouped by day to reuse verified day views, but results remain in input order. An item failure does not discard other results; cancellation/deadline errors terminate the job.

`search_mail` scrolls Outlook’s result list until it has `limit` identifiable rows or the list stops rendering more. It still does not call a mail API, and it skips rows whose sender or subject it cannot see, so a broad query can return fewer hits than `limit`. Narrow queries stay faster and less ambiguous. A row can represent a conversation, not an individual message. When a conversation shows several expanded messages, `read_mail` returns each visible body separated by `---`, not necessarily the entire thread. Images are blocked; their alt text appears as `(image: …)`. Bodies are capped at `max_body_chars` (default and maximum 12,000 characters); `truncated` reports omitted content.

A result row with an Outlook item id (`data-item-id`) or conversation id (`data-convid`), including on a descendant of the row, gets a stable `ow1.` message ID. That ID embeds the search query plus the Outlook id, sender, subject, time and preview, so `read_mail` and `read_attachment` can run the search again and match the same row after a server restart or after the 30-minute in-memory cache expires. Treat the token as opaque mailbox metadata. Rows without either attribute still get a random UUID. Those temporary IDs, including ones already issued, keep working until they expire or the process restarts. A stable ID that no longer matches a single row returns `MESSAGE_CHANGED`.

`search_events` reads Outlook's day view from `start` through `end`, end-inclusive, at most 14 days. If you omit the dates, it uses the day the calendar opens. `query` filters the rendered title and location; it is not Outlook search syntax and does not look outside that range. Scanning stops after a day provides enough matches for `limit`; `scanned_start`, `scanned_end`, and `stopped_early` identify the actual scanned range, and `truncated` also covers omitted matches on the last scanned day. Results remain date-first. `event_id` expires after 30 minutes, as temporary mail IDs do. `read_event` opens that event's details peek. Notes use the same Markdown rules as mail: images become alt text, links are not followed, and the body is capped at 12,000 characters.

`mail_status` reports **last observed** state, not a live authentication check. It does not start a browser, sign in, or query Outlook. Use the local control command `npm run status -- --check` for a live mailbox check.

A calendar request confirms the signed-in mailbox, then uses a separate tab so the mail list stays put. The tools never create, edit, delete or respond to events. A single click opens the peek. Accept, Tentative, Decline, Edit, Delete and Save are not clicked. If Outlook opens the editor instead of the peek, the adapter closes it and discards unsaved changes. Only calendars already selected in Outlook are visible, and only the events drawn for each requested day. The date is read from a machine-readable attribute when Outlook provides one, otherwise from an English heading such as "12 November 2026". Times are returned as Outlook renders them, including a datetime attribute when one is present. An aria-label without structured fields is parsed on commas, so a title that itself contains a comma can be cut short.

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

`npm run login` first attempts unattended sign-in when `OUTLOOK_USERNAME`, `OUTLOOK_PASSWORD` or `OUTLOOK_PASSWORD_FILE` is configured in `.env`. It uses the existing session or the configured password and exits with an error if Microsoft needs MFA or another manual step. Use `npm run login -- --interactive` to open a sign-in window, or `npm run login -- --non-interactive` to check an existing session without a desktop. With no credentials configured, `npm run login` opens a sign-in window. Complete Microsoft sign-in and MFA and choose **Stay signed in** when asked: that persistent sign-in is what lets the server run headless afterwards. The window closes by itself once Outlook loads. For Outlook.com, put `OUTLOOK_URL=https://outlook.live.com/mail/` in `.env` (the default is `https://outlook.office.com/mail/`).

Opening a message can mark it as read. Afterwards the adapter clicks a visible **Mark as unread** button outside the message body when the row was labeled unread. That click does not change Outlook’s settings, and it does nothing when the button is absent, so disable automatic marking as read where available and verify with an unread test message before using the server on real mail. Use a dedicated profile: logout deletes its directory. The profile lives in `~/.local/state/outlook-mcp/chromium` (`OUTLOOK_PROFILE`). Keep it private.

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
| `npm run login` | Attempts unattended sign-in with configured credentials; `-- --interactive` opens a visible sign-in window, even when the server runs as a service. Mail requests answer `AUTH_REQUIRED` until it finishes. Times out after 10 minutes. |
| `npm run logout` | Closes the browser and deletes the browser profile, removing every saved Microsoft sign-in. |
| `npm run status` | Shows sign-in state, whether the browser is running, the last success or error, and whether the password file is valid. Exits with an error if `OUTLOOK_PASSWORD_FILE` is set but invalid. Add `-- --check` to load Outlook and test the session now. |

When Microsoft ends the session, the server first tries to resume it without a secret: it picks the remembered account in the account picker and accepts "Stay signed in?". If Microsoft still wants a password or MFA, mail requests return `AUTH_REQUIRED`. The message, the journal line, and the desktop notification say which case it is: the password file is misconfigured, the Microsoft session expired and needs `npm run login`, a login window is already open, or another case such as a cooldown after a failed automatic attempt. `OUTLOOK_NOTIFY=0` turns notifications off. Login and logout are only available on the local socket, never on the MCP port, so a tailnet client cannot trigger them. `mail_status` reports the same sign-in facts over MCP and cannot sign in or sign out.

Federated sign-in pages outside Microsoft (for example an organisation's ADFS) are blocked in headless mode. Add their origins to `OUTLOOK_SIGNIN_ORIGINS`, comma-separated.

## Keep a visible browser open instead

If Microsoft will not keep you signed in between browser restarts, set `OUTLOOK_INTERACTIVE=1` in `.env` and run `npm start` in a terminal (the service forces headless mode). The server opens a window; complete sign-in there and leave it open while using the MCP. If you close it or restart the server, sign in again.

## Automatically sign back in

For accounts that use a password without MFA, put the credentials in `.env` (keep it mode `0600`) and restart the server:

```dotenv
OUTLOOK_USERNAME=your-microsoft-sign-in-address
OUTLOOK_PASSWORD=your-microsoft-password
```

To keep the password out of `.env`, set `OUTLOOK_PASSWORD_FILE=/absolute/path/to/a/private/outlook-password` instead. Use one or the other: setting both stops the server at startup. The file must have mode `0600` (owner-only modes such as `0400` are fine) and hold only the password on its first line. Group or world permissions are rejected. Either way the password is plaintext readable by your OS account, so keep it out of syncing and backups. Chromium and `notify-send` are started without `OUTLOOK_PASSWORD` or `MCP_TOKEN` in their environment, and the password is never logged or returned by `mail_status`.

When `OUTLOOK_PASSWORD_FILE` is set, the server checks it before listening and exits if it is invalid. `npm run status` and `npm run status -- --check` check it too. A failure names the configured path and the problem (missing, not a regular file, unreadable, empty, or loose permissions). A missing file also lists similarly named files in the same directory as hints, but nothing substitutes a different file.

How the server handles a sign-out:

- Outlook often loads on its own origin and redirects to Microsoft sign-in a second or two later. The server keeps watching for that redirect instead of giving up.
- It first resumes without a secret: it picks the remembered account, accepts **Stay signed in?** (ticking "Don't show this again") and clicks **Ask later** or **Skip for now** on optional security prompts. Only then does it type the configured username and password.
- When Outlook shows its own "session expired" prompt, the server reloads the mailbox to start the Microsoft sign-in flow.
- When the session ends during a request, including in the calendar tab, the server signs back in once and repeats the request before reporting `AUTH_REQUIRED`.
- When Microsoft rejects the password or username, or reports the account as locked, automatic sign-in pauses for 6 hours so repeated attempts cannot lock the account. When MFA or another verification step is required, it pauses for 15 minutes. Any other failure retries after 2 minutes. `npm run login` clears the pause at any time.
- Automatic recoveries are logged to the journal (`signed back in automatically…`). A desktop notification appears only when a person needs to act.

With credentials configured, `npm run login` attempts this unattended sign-in. Use `npm run login -- --interactive` for MFA.

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

Tests use Playwright's Chromium, falling back to `/usr/bin/chromium` on Linux (override with `CHROMIUM_PATH`). The tests exercise Chromium fixtures (including expired sessions, the account picker, automatic sign-in, crashes, hangs, network loss and sign-in/sign-out), password-file checks, the authenticated MCP handshake, `mail_status`, and the control socket. The fixture does not validate selectors against your signed-in Outlook layout. After sign-in, search for a known unread email, call `read_mail` with its returned ID, and verify its subject, body and unread state in Outlook.

Temporary search result IDs expire after 30 minutes or on server restart. Stable `ow1.` IDs do not. Error codes:

- `UNKNOWN_MESSAGE`: the ID is unknown, malformed, or a temporary ID that expired. Search again.
- `MESSAGE_CHANGED`: the result moved or no longer matches one row. Search again.
- `AUTH_REQUIRED`: sign-in is required. The message says which: the password file is misconfigured (path and reason, plus similar filenames when it is missing), the Microsoft session expired (`npm run login`), a login window is already open, or another case such as a cooldown after automatic sign-in failed.
- `UNKNOWN_EVENT` / `EVENT_CHANGED`: the event ID expired or that event is no longer the same item. Search again.
- `INVALID_RANGE`: the dates, query or limit cannot be used. Ranges are real `YYYY-MM-DD` dates of at most 14 days.
- `SIGNED_OUT`: `npm run logout` is running.
- `NETWORK_ERROR`: Outlook could not be reached.
- `TIMEOUT`: the three-minute enqueue deadline elapsed. Queued work is skipped; a timed-out active browser operation resets the browser. Retry.
- `CANCELLED`: the MCP request was cancelled or disconnected. Queued work is skipped; active work stops at checkpoints.
- `STORAGE_FULL`: retained attachment storage is full of unexpired, unserved files. Retry after links expire.
- `ATTACHMENT_TOO_LARGE`: the visible/actual file size or retained-byte budget is exceeded.
- `BROWSER_RESTARTED`: the browser crashed or its window closed. The next request starts a new one.
- `PROFILE_LOCKED`: another browser (an old `npm start`, or a direct `npm run login`) is using the profile.
- `UI_CHANGED`: the adapter could not identify a search region, row or reading pane. Outlook error pages get one automatic reload first.
- `BROWSER_ERROR`: anything else.

For `TIMEOUT`, `UI_CHANGED` and `BROWSER_ERROR`, the server prints the failing Playwright step to stderr. Outlook's DOM differs between accounts and changes over time. The mail selectors are in `src/outlook.ts`, and the fixture in `test/mail.test.mjs` mirrors the layout they were checked against (`outlook.cloud.microsoft`, November 2026). Calendar selectors are in `src/calendar.ts`. The sanitized pages in `test/fixtures/calendar/` mirror the day view and event peek they target. Neither fixture is a live mailbox.

Email fields are untrusted source data. The worker blocks images during mailbox operations and does not follow email links. It downloads an attachment only when `read_attachment` asks for it, never opens or runs it, and keeps a temporary copy for up to 30 minutes for download. Attachment content is untrusted in the same way as email bodies. The server does not store email bodies. Your MCP client and model provider may receive and retain returned content according to their own settings.

## Development and security

See [CONTRIBUTING.md](CONTRIBUTING.md) for development checks and sanitized bug
reports, [SECURITY.md](SECURITY.md) for the security model, and
[.env.example](.env.example) for configuration settings. GitHub CI runs the fixture
suite on Node.js 22, 24 and 26. These checks do not validate a live mailbox.

## License

[ISC](LICENSE).


### Temporary attachment downloads

The MCP process also serves GET-only attachment links on its `tailscale0` IPv4 address,
port 8766 (override with `OUTLOOK_DOWNLOAD_PORT`). It refuses to start the attachment
listener on any other interface. `OUTLOOK_DOWNLOAD_HOSTNAME` overrides the URL hostname;
the default is the machine hostname plus `.tail4832d.ts.net`.

Each download has a fresh 256-bit random token. Anyone on the tailnet with the link can
download it until its 30-minute expiry; treat links as secrets. Expired links immediately
return 404. A cleanup pass deletes expired copies every 30 seconds. Retained files total at most
100 MiB (`OUTLOOK_ATTACHMENT_RETAINED_MB` overrides this). When space is needed, expired files
and then the oldest already-served files are evicted; served links may therefore disappear before
their advertised expiry. Unserved, unexpired files are not evicted: new downloads get `STORAGE_FULL`
if no room remains. In-progress GETs are protected from eviction. Files have mode 0600
inside a private temporary directory. Request paths are matched against an in-memory map,
never resolved as filesystem paths; there is no directory listing. All attachment types
return text metadata, so MCP hosts need not handle embedded binary resources.

The existing systemd service supervises both listeners in one process. Stopping or
restarting invalidates all links and deletes temporary copies. The installed service's
`PrivateTmp=yes` also removes its private temporary files when the service stops.
Mail operations and their read-only annotations are unchanged.
