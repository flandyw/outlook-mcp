# Security

This project automates a signed-in browser; its dedicated profile grants access
to your mailbox. Keep the profile, bearer token, optional password file, and `.env`
private. Only use a profile dedicated to this application. Do not point
`OUTLOOK_PROFILE` at your everyday browser profile: logout deletes that directory.

The HTTP endpoint binds to loopback and requires a bearer token. Remote access
requires a protected tunnel or reverse proxy and an appropriate access policy.
The local control socket permits sign-in and profile deletion for your OS account.
Startup, `npm run status`, and the read-only `mail_status` tool report whether
the password file is configured and valid, including its path. They do not
include the file contents. `mail_status` cannot sign in, sign out, or read mail.

The tools do not send, reply, delete, or move mail, and they do not expose a
flag action. Opening a message can still mark it as read. When the list row was
unread and Outlook shows exactly one visible button named "Mark as unread"
outside the message body, the adapter clicks that button after a successful
read or attachment download. Controls inside the message body are never
clicked. If the button is absent, the message may stay read. Returned email is
untrusted input and may contain prompt injection. Clients must treat it as data,
not instructions. Returned content may be retained by the client or model provider.

Do not put credentials, private mailbox content, or exploit details in public
issues. Use the repository's private vulnerability reporting feature if enabled.
Otherwise contact the repository owner privately before disclosing details.
Security fixes target the latest version on `main`; older versions have no
separate maintenance commitment.
