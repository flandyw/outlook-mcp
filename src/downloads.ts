import { createServer } from "node:http";
import { randomBytes } from "node:crypto";
import { mkdtemp, writeFile, rm } from "node:fs/promises";
import { createReadStream } from "node:fs";
import { tmpdir, networkInterfaces } from "node:os";
import { join } from "node:path";
import { pipeline } from "node:stream/promises";

/** Exact URL lookup only: request paths are never used as filesystem paths. */
export async function startDownloads({ host, port = 8766, hostname, ttlMs = 30 * 60_000,
  allowLoopbackForTest = false }: { host: string; port?: number; hostname: string; ttlMs?: number; allowLoopbackForTest?: boolean }) {
  const addresses = networkInterfaces().tailscale0 ?? [];
  if (!(allowLoopbackForTest && host === "127.0.0.1") &&
      !addresses.some(address => address.address === host && address.family === "IPv4")) {
    throw new Error("Attachment server must bind to an IPv4 address on tailscale0.");
  }
  const dir = await mkdtemp(join(tmpdir(), "outlook-attachments-"));
  const files = new Map<string, { path: string; name: string; mime: string; size: number; expires: number }>();
  async function cleanup() {
    await Promise.all([...files].filter(([, file]) => file.expires <= Date.now()).map(async ([key, file]) => {
      await rm(file.path, { force: true });
      files.delete(key);
    }));
  }
  const server = createServer(async (req, res) => {
    res.setHeader("Cache-Control", "no-store");
    res.setHeader("X-Content-Type-Options", "nosniff");
    if (req.method !== "GET") return void res.writeHead(405, { Allow: "GET" }).end();
    const file = files.get(req.url ?? "");
    if (!file || file.expires <= Date.now()) return void res.writeHead(404).end();
    res.writeHead(200, { "Content-Type": file.mime, "Content-Length": file.size,
      "Content-Disposition": "attachment; filename*=UTF-8''" + encodeURIComponent(file.name) });
    await pipeline(createReadStream(file.path), res).catch(() => res.destroy());
  });
  try {
    await new Promise<void>((resolve, reject) => server.once("error", reject).listen(port, host, resolve));
  } catch (error) {
    await rm(dir, { recursive: true, force: true });
    throw error;
  }
  const timer = setInterval(() => void cleanup().catch(error => console.error("Attachment cleanup failed", error)), 30_000);
  timer.unref();
  return {
    async save(name: string, data: Buffer, mime: string) {
      const token = randomBytes(32).toString("hex");
      const safeName = name.replace(/[\\/\x00-\x1f\x7f]/g, "_").slice(0, 200) || "attachment";
      const key = "/" + token + "/" + encodeURIComponent(safeName);
      const path = join(dir, token);
      await writeFile(path, data, { mode: 0o600, flag: "wx" });
      const expires = Date.now() + ttlMs;
      files.set(key, { path, name: safeName, mime, size: data.length, expires });
      const address = server.address();
      return { download_url: `http://${hostname}:${typeof address === "object" && address ? address.port : port}${key}`,
        expires_at: new Date(expires).toISOString() };
    },
    cleanup,
    directory: dir,
    async close() {
      clearInterval(timer);
      server.closeAllConnections();
      await new Promise<void>(resolve => server.close(() => resolve()));
      await rm(dir, { recursive: true, force: true });
    },
  };
}
