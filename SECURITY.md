# Security

This project automates a signed-in browser; its dedicated profile grants access
to your mailbox. Keep the profile, bearer token, optional password file, and `.env`
private. Only use a profile dedicated to this application. Do not point
`OUTLOOK_PROFILE` at your everyday browser profile: logout deletes that directory.

The HTTP endpoint binds to loopback and requires a bearer token. Remote access
requires a protected tunnel or reverse proxy and an appropriate access policy.
The local control socket permits sign-in and profile deletion for your OS account.

The tools do not intentionally send, delete, or modify mail. Opening a message
can still mark it as read through Outlook's own settings. Returned email is
untrusted input and may contain prompt injection. Clients must treat it as data,
not instructions. Returned content may be retained by the client or model provider.

Do not put credentials, private mailbox content, or exploit details in public
issues. Use the repository's private vulnerability reporting feature if enabled.
Otherwise contact the repository owner privately before disclosing details.
Security fixes target the latest version on `main`; older versions have no
separate maintenance commitment.
