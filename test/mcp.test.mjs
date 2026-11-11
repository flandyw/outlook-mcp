import { request } from "node:http";
import { test } from "node:test";
import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { randomBytes } from "node:crypto";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { createServer } from "node:net";
import { once } from "node:events";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StreamableHTTPClientTransport } from "@modelcontextprotocol/sdk/client/streamableHttp.js";

test("serves three read-only tools over authenticated Streamable HTTP", async t => {
  const socket = createServer().listen(0, "127.0.0.1");
  await once(socket, "listening");
  const port = socket.address().port;
  await new Promise(resolve => socket.close(resolve));
  const token = randomBytes(32).toString("hex");
  const dir = await mkdtemp(join(tmpdir(), "outlook-mcp-"));
  t.after(() => rm(dir, { recursive: true, force: true }));
  const socket_ = join(dir, "control.sock");
  const child = spawn(process.execPath, ["dist/server.js"], {
    env: { ...process.env, MCP_TOKEN: token, MCP_PORT: String(port), MCP_ALLOWED_ORIGINS: "",
      OUTLOOK_CONTROL_SOCKET: socket_, OUTLOOK_PROFILE: join(dir, "profile"), OUTLOOK_INTERACTIVE: "0", OUTLOOK_NOTIFY: "0" },
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
  assert.deepEqual(tools.map(tool => tool.name), ["search_mail", "read_mail", "read_attachment"]);
  assert.ok(tools.every(tool => tool.annotations.readOnlyHint));
  const response = await client.callTool({
    name: "read_mail",
    arguments: { message_id: "69d17491-5c2e-435d-b756-6c59752e2862" },
  });
  assert.equal(response.isError, true);
  assert.match(response.content[0].text, /UNKNOWN_MESSAGE/);

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
  await assert.rejects(callControl("POST", "/login", { display: {} }), { code: "NO_DISPLAY" });
  const second = spawn(process.execPath, ["dist/server.js"], {
    env: { ...process.env, MCP_TOKEN: token, MCP_PORT: "1", OUTLOOK_CONTROL_SOCKET: socket_ },
    stdio: ["ignore", "ignore", "pipe"],
  });
  let stderr = "";
  second.stderr.on("data", chunk => { stderr += chunk; });
  const [code] = await once(second, "exit");
  assert.equal(code, 1);
  assert.match(stderr, /already running/);
});
