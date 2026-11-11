# Contributing

Use Node.js 22.9 or newer and npm. Linux is the primary target; the service
installer requires systemd and the control commands use Unix sockets. Native
Windows is not supported.

```sh
npm ci
npx playwright install --with-deps chromium
npm run check
npm test
```

Tests use local browser fixtures and temporary directories, with no Microsoft
account or real mailbox required. Do not load your personal `.env` when running
tests. Live Outlook behavior needs a separate manual check because Microsoft's
layout changes independently of these fixtures.

Keep changes focused, add regression coverage for behavioral changes, and describe
what you tested in your pull request. Avoid unrelated dependency or lockfile changes.

Never include tokens, passwords, browser profiles, real email content, screenshots
of mailboxes, or logs containing personal information in commits or public issues.
For selector bugs, use a minimal sanitized HTML fixture and describe the layout.
See [SECURITY.md](SECURITY.md) for security concerns.
