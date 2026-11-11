import { request } from "node:http";
import { homedir } from "node:os";
import { join } from "node:path";

/**
 * Local control socket for login, logout and status. It is a Unix socket in a 0700 directory,
 * so only this OS account can use it, and it never shares the (possibly tailnet-exposed) MCP port.
 */
export function controlSocketPath() {
  if (process.env.OUTLOOK_CONTROL_SOCKET) return process.env.OUTLOOK_CONTROL_SOCKET;
  const base = process.env.XDG_RUNTIME_DIR ?? join(homedir(), ".local/state");
  return join(base, "outlook-mcp", "control.sock");
}

export class ServerNotRunning extends Error {}

export function callControl<T = Record<string, unknown>>(method: "GET" | "POST", path: string, body?: unknown, timeoutMs = 15 * 60_000) {
  return new Promise<T>((resolve, reject) => {
    const payload = body === undefined ? undefined : JSON.stringify(body);
    const req = request({
      socketPath: controlSocketPath(), method, path,
      headers: payload ? { "Content-Type": "application/json", "Content-Length": Buffer.byteLength(payload) } : {},
      timeout: timeoutMs,
    }, res => {
      let text = "";
      res.setEncoding("utf8");
      res.on("data", chunk => { text += chunk; });
      res.on("end", () => {
        let data: unknown;
        try {
          data = JSON.parse(text);
        } catch {
          return reject(new Error("Unexpected reply from the control socket."));
        }
        if (res.statusCode && res.statusCode >= 400) {
          const error = data as { code?: string; message?: string };
          return reject(Object.assign(new Error(error.message ?? "Control request failed."), { code: error.code }));
        }
        resolve(data as T);
      });
    });
    req.on("timeout", () => req.destroy(new Error("The server did not answer in time.")));
    req.on("error", (error: NodeJS.ErrnoException) => {
      reject(error.code === "ENOENT" || error.code === "ECONNREFUSED" ? new ServerNotRunning() : error);
    });
    req.end(payload);
  });
}
