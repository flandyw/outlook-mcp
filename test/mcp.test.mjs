import { request } from "node:http";
import { test } from "node:test";
import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { randomBytes } from "node:crypto";
import { chmod, mkdtemp, readdir, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { createServer } from "node:net";
import { once } from "node:events";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StreamableHTTPClientTransport } from "@modelcontextprotocol/sdk/client/streamableHttp.js";

test("serves read-only mail and calendar tools over authenticated Streamable HTTP", async t => {
  const socket = createServer().listen(0, "127.0.0.1");
  await once(socket, "listening");
  const port = socket.address().port;
  await new Promise(resolve => socket.close(resolve));
  const token = randomBytes(32).toString("hex");
  const dir = await mkdtemp(join(tmpdir(), "outlook-mcp-"));
  t.after(() => rm(dir, { recursive: true, force: true }));
  const socket_ = join(dir, "control.sock");
  const password = join(dir, "password");
  const canary = "mcp-canary-do-not-print";
  await writeFile(password, canary + "\n");
  await chmod(password, 0o600);
  const child = spawn(process.execPath, ["--import", "./test/fixtures/tailscale-interface.mjs", "--import", "./test/fixtures/mcp-tools.mjs", "dist/server.js"], {
    env: { ...process.env, MCP_TOKEN: token, MCP_PORT: String(port), MCP_ALLOWED_ORIGINS: "", OUTLOOK_DOWNLOAD_PORT: String(port + 1),
      OUTLOOK_CONTROL_SOCKET: socket_, OUTLOOK_PROFILE: join(dir, "profile"), OUTLOOK_INTERACTIVE: "0", OUTLOOK_NOTIFY: "0",
      OUTLOOK_PASSWORD_FILE: password, OUTLOOK_DOWNLOAD_HOSTNAME: "127.0.0.1" },
    stdio: ["ignore", "pipe", "inherit"],
  });
  t.after(() => child.kill());
  await Promise.race([
    once(child.stdout, "data"),
    once(child, "exit").then(() => { throw new Error("MCP server exited before startup"); }),
  ]);
  const url = new URL("http://127.0.0.1:" + port + "/mcp");
  const headers = { Authorization: "Bearer " + token };
  assert.equal((await fetch(url, { method: "POST" })).status, 401);
  assert.equal((await fetch(url, {
    method: "POST", headers: { ...headers, Origin: "https://example.com" },
  })).status, 403);
  const client = new Client({ name: "test", version: "1.0.0" });
  t.after(() => client.close());
  await client.connect(new StreamableHTTPClientTransport(url, { requestInit: { headers } }));
  const { tools } = await client.listTools();
  assert.deepEqual(tools.map(tool => tool.name), ["list_mail", "search_mail", "read_mail", "read_mails", "read_attachment", "mail_status", "search_events", "read_event", "read_events"]);
  assert.ok(tools.every(tool => tool.annotations.readOnlyHint));
  assert.ok(tools.every(tool => tool.annotations.destructiveHint === false));
  for (const tool of tools) {
    assert.ok(tool.title);
    assert.equal(tool.outputSchema.type, "object");
    assert.ok(Object.keys(tool.outputSchema.properties).length);
    for (const param of Object.values(tool.inputSchema.properties)) assert.ok(param.description);
    if (tool.name !== "mail_status") assert.notEqual(tool.annotations.idempotentHint, true);
  }
  const readMail = tools.find(tool => tool.name === "read_mail");
  assert.equal(readMail.inputSchema.properties.max_body_chars.default, 12000);
  assert.equal(readMail.inputSchema.properties.max_body_chars.maximum, 12000);
  assert.equal(tools.find(tool => tool.name === "read_mails").inputSchema.properties.message_ids.maxItems, 10);
  assert.equal(tools.find(tool => tool.name === "read_events").inputSchema.properties.event_ids.minItems, 1);
  const statusTool = tools.find(tool => tool.name === "mail_status");
  assert.equal(statusTool.annotations.destructiveHint, false);
  assert.equal(statusTool.annotations.openWorldHint, false);
  const reported = await client.callTool({ name: "mail_status", arguments: {} });
  assert.ok(!reported.isError);
  const diagnostics = reported.structuredContent;
  assert.match(reported.content[0].text, /Last observed/);
  assert.doesNotMatch(reported.content[0].text, /^\s*\{/);
  assert.equal(diagnostics.auth, "unknown");
  assert.equal(diagnostics.browser, "stopped");
  assert.equal(diagnostics.busy, null);
  assert.equal(diagnostics.last_success_at, null);
  assert.equal(diagnostics.last_error, null);
  assert.equal(diagnostics.password_file_configured, true);
  assert.equal(diagnostics.password_file_valid, true);
  assert.equal(diagnostics.password_file_path, password);
  assert.equal(diagnostics.password_file_problem, null);
  assert.deepEqual(diagnostics.password_file_hints, []);
  assert.equal(diagnostics.profile, undefined);
  assert.equal(JSON.stringify(diagnostics).includes(canary), false);
  assert.deepEqual(tools.map(tool => tool.name).filter(name => /login|logout|send|reply|delete|move|flag/i.test(name)), []);
  const searchEvents = tools.find(tool => tool.name === "search_events");
  assert.ok(searchEvents.inputSchema.properties.start);
  assert.ok(searchEvents.inputSchema.properties.query);
  const response = await client.callTool({
    name: "read_mail",
    arguments: { message_id: "69d17491-5c2e-435d-b756-6c59752e2862" },
  });
  assert.equal(response.isError, true);
  assert.match(response.content[0].text, /UNKNOWN_MESSAGE/);
  assert.deepEqual(response.structuredContent, { code: "UNKNOWN_MESSAGE", retryable: false, message: "Run search_mail again to get a fresh message ID." });
  const event = await client.callTool({
    name: "read_event",
    arguments: { event_id: "69d17491-5c2e-435d-b756-6c59752e2862" },
  });
  assert.equal(event.isError, true);
  assert.match(event.content[0].text, /UNKNOWN_EVENT/);
  assert.equal(event.structuredContent.code, "UNKNOWN_EVENT");
  assert.equal(event.structuredContent.retryable, false);
  for (const [name, args] of [["read_mails", { message_ids: [] }], ["read_events", { event_ids: [] }],
    ["read_mail", { message_id: "69d17491-5c2e-435d-b756-6c59752e2862", max_body_chars: 12001 }]]) {
    assert.equal((await client.callTool({ name, arguments: args })).isError, true);
  }

  // Malformed JSON and non-object payloads must not become internal errors.
  for (const payload of ["{", "null", "[]", "42"]) {
    const response = await new Promise((resolve, reject) => {
      const req = request({ socketPath: socket_, method: "POST", path: "/login" }, res => {
        let body = "";
        res.setEncoding("utf8");
        res.on("data", chunk => { body += chunk; });
        res.on("end", () => resolve({ status: res.statusCode, body: JSON.parse(body) }));
      });
      req.on("error", reject);
      req.end(payload);
    });
    assert.equal(response.status, 400);
    assert.equal(response.body.code, "INVALID_REQUEST");
  }

  // The control socket answers status and refuses a second server.
  process.env.OUTLOOK_CONTROL_SOCKET = socket_;
  const { callControl } = await import("../dist/control.js");
  const status = await callControl("GET", "/status");
  assert.equal(status.auth, "unknown");
  assert.equal(status.browser, "stopped");
  assert.equal(status.password_file.configured, true);
  assert.equal(status.password_file.valid, true);
  assert.equal(status.password_file.path, password);
  assert.equal(JSON.stringify(status).includes(canary), false);
  await chmod(password, 0o644);
  const after = await client.callTool({ name: "mail_status", arguments: {} });
  const loosened = after.structuredContent;
  assert.equal(loosened.browser, "stopped");
  assert.equal(loosened.password_file_valid, false);
  assert.equal(loosened.password_file_problem, "loose_permissions");
  assert.match(loosened.password_file_path, new RegExp(password.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")));
  assert.equal(JSON.stringify(loosened).includes(canary), false);
  const controlAfter = await callControl("GET", "/status");
  assert.equal(controlAfter.password_file.problem, "loose_permissions");
  assert.equal(JSON.stringify(controlAfter).includes(canary), false);
  await chmod(password, 0o600);
  await assert.rejects(callControl("POST", "/login", { display: {} }), { code: "NO_DISPLAY" });
  const found = await client.callTool({ name: "search_mail", arguments: { query: "fixture" } });
  const listed = await client.callTool({ name: "list_mail", arguments: {} });
  assert.equal(listed.structuredContent.folder, "inbox");
  assert.equal(listed.structuredContent.messages[0].unread, true);
  assert.equal((await client.callTool({ name: "list_mail", arguments: { folder: "../calendar" } })).isError, true);
  assert.ok(!found.isError);
  const message_id = found.structuredContent.messages[0].message_id;
  const read = await client.callTool({ name: "read_mail", arguments: { message_id, max_body_chars: 5 } });
  assert.equal(read.structuredContent.body_markdown, "Fixtu");
  assert.equal(read.structuredContent.truncated, true);
  const batch = await client.callTool({ name: "read_mails", arguments: { message_ids: [message_id, "69d17491-5c2e-435d-b756-6c59752e2862"] } });
  assert.ok(!batch.isError);
  assert.equal(batch.structuredContent.results[0].data.subject, "Fixture mail");
  assert.equal(batch.structuredContent.results[1].error.code, "UNKNOWN_MESSAGE");
  const file = await client.callTool({ name: "read_attachment", arguments: { message_id, attachment: 0 } });
  assert.ok(!file.isError);
  const resource = file.content.find(item => item.type === "resource_link");
  assert.equal(resource.uri, file.structuredContent.download_url);
  assert.equal(resource.name, "fixture.txt");
  assert.equal(await (await fetch(resource.uri)).text(), "fixture");
  const days = await client.callTool({ name: "search_events", arguments: {} });
  assert.ok(!days.isError);
  const event_id = days.structuredContent.events[0].event_id;
  assert.ok(!(await client.callTool({ name: "read_event", arguments: { event_id } })).isError);
  const events = await client.callTool({ name: "read_events", arguments: { event_ids: [event_id, "69d17491-5c2e-435d-b756-6c59752e2862"] } });
  assert.ok(!events.isError);
  assert.equal(events.structuredContent.results[0].data.title, "Fixture event");
  assert.equal(events.structuredContent.results[1].error.code, "UNKNOWN_EVENT");
  const controller = new AbortController();
  const waitingForCancel = once(child.stdout, "data");
  const cancelled = client.callTool({ name: "search_mail", arguments: { query: "wait" } }, undefined, { signal: controller.signal });
  await waitingForCancel;
  controller.abort();
  await assert.rejects(cancelled);
  assert.ok(!(await client.callTool({ name: "search_mail", arguments: { query: "after cancellation" } }, undefined, { timeout: 3000 })).isError,
    "SDK cancellation notification must stop work across stateless POSTs, even without disconnecting");
  const held = [];
  for (let index = 0; index < 8; index++) {
    const enqueued = once(child.stdout, "data");
    held.push(new Promise((resolve, reject) => {
      const req = request(url, { method: "POST", headers: { ...headers, "Content-Type": "application/json", Accept: "application/json, text/event-stream" } }, res => {
        let body = "";
        res.on("data", chunk => { body += chunk; });
        res.on("end", () => resolve(JSON.parse(body)));
      });
      req.on("error", reject);
      req.end(JSON.stringify({ jsonrpc: "2.0", id: 1000 + index, method: "tools/call", params: { name: "search_mail", arguments: { query: "hold" } } }));
    }));
    assert.match(String((await enqueued)[0]), /fixture enqueued/);
  }
  for (let index = 0; index < 8; index++) {
    const response = await fetch(url, { method: "POST", headers: { ...headers, "Content-Type": "application/json", Accept: "application/json, text/event-stream" },
      body: JSON.stringify({ jsonrpc: "2.0", method: "notifications/cancelled", params: { requestId: 1000 + index } }) });
    assert.equal(response.status, 202, "cancellation must remain available when all eight operation slots are occupied");
  }
  for (const response of await Promise.all(held)) assert.equal(response.result.structuredContent.code, "CANCELLED");
  for (let index = 0; index < 9; index++) {
    const waiting = once(child.stdout, "data");
    const disconnected = request(url, { method: "POST", headers: { ...headers, "Content-Type": "application/json", Accept: "application/json, text/event-stream" } });
    disconnected.on("error", () => {});
    disconnected.end(JSON.stringify({ jsonrpc: "2.0", id: 100 + index, method: "tools/call", params: { name: "search_mail", arguments: { query: "wait" } } }));
    assert.match(String((await waiting)[0]), /fixture waiting/);
    disconnected.destroy();
    const after = await client.callTool({ name: "search_mail", arguments: { query: "after disconnect" } }, undefined, { timeout: 3000 });
    assert.ok(!after.isError, "disconnect must abort handler extra.signal and release its HTTP slot");
  }
  const second = spawn(process.execPath, ["--import", "./test/fixtures/tailscale-interface.mjs", "dist/server.js"], {
    env: { ...process.env, MCP_TOKEN: token, MCP_PORT: "1", OUTLOOK_CONTROL_SOCKET: socket_, OUTLOOK_PASSWORD_FILE: "" },
    stdio: ["ignore", "ignore", "pipe"],
  });
  let stderr = "";
  second.stderr.on("data", chunk => { stderr += chunk; });
  const [code] = await once(second, "exit");
  assert.equal(code, 1);
  assert.match(stderr, /already running/);
});

test("refuses to start when OUTLOOK_PASSWORD_FILE is invalid", async t => {
  const dir = await mkdtemp(join(tmpdir(), "outlook-mcp-password-"));
  t.after(async () => {
    for (const name of await readdir(dir).catch(() => [])) await chmod(join(dir, name), 0o700).catch(() => {});
    await rm(dir, { recursive: true, force: true });
  });
  const token = randomBytes(32).toString("hex");
  const canary = "startup-canary-do-not-print";
  const notesCanary = "notes-canary-do-not-print";
  const configured = join(dir, "alpha-token-value");
  await writeFile(join(dir, "alpha-token-value-bak"), canary + "\n");
  await chmod(join(dir, "alpha-token-value-bak"), 0o600);
  await writeFile(join(dir, "notes.txt"), notesCanary + "\n");
  await chmod(join(dir, "notes.txt"), 0o600);
  const loose = join(dir, "loose");
  await writeFile(loose, canary + "\n");
  await chmod(loose, 0o644);

  async function rejected(file, pattern) {
    const child = spawn(process.execPath, ["dist/server.js"], {
      env: {
        PATH: process.env.PATH, HOME: process.env.HOME, TMPDIR: process.env.TMPDIR,
        MCP_TOKEN: token, MCP_PORT: "9", OUTLOOK_PASSWORD_FILE: file,
        OUTLOOK_PROFILE: join(dir, "profile"), OUTLOOK_CONTROL_SOCKET: join(dir, "control.sock"),
        OUTLOOK_NOTIFY: "0", OUTLOOK_ATTACHMENTS: "0",
      },
      stdio: ["ignore", "pipe", "pipe"],
    });
    let stdout = "";
    let stderr = "";
    child.stdout.on("data", chunk => { stdout += chunk; });
    child.stderr.on("data", chunk => { stderr += chunk; });
    const [code] = await Promise.race([
      once(child, "exit"),
      new Promise((_, reject) => setTimeout(() => {
        child.kill();
        reject(new Error("server did not exit\n" + stderr));
      }, 8000)),
    ]);
    assert.equal(code, 1);
    assert.equal(stdout, "");
    assert.match(stderr, /outlook-mcp: Password file misconfigured/);
    assert.match(stderr, pattern);
    assert.equal(stderr.includes(canary), false);
    assert.equal(stderr.includes(notesCanary), false);
    assert.doesNotMatch(stdout + stderr, /listening/i);
    return stderr;
  }

  const missing = await rejected(configured, /is missing/);
  assert.equal(missing.includes(join(dir, "alpha-token-value-bak")), true);
  assert.match(missing, /not used automatically/);
  assert.equal(missing.includes(join(dir, "notes.txt")), false);
  const looseError = await rejected(loose, /loose permissions/);
  assert.match(looseError, /0644/);
});
