import { createServer } from "node:http";
import { randomBytes } from "node:crypto";
import { mkdtemp, writeFile, rm, rename, copyFile, chmod, stat } from "node:fs/promises";
import { createReadStream } from "node:fs";
import { tmpdir, networkInterfaces } from "node:os";
import { join } from "node:path";
import { pipeline } from "node:stream/promises";
import { MailError } from "./outlook.js";

/** Exact URL lookup only: request paths are never used as filesystem paths. */
export async function startDownloads({ host, port = 8766, hostname, ttlMs = 30 * 60_000,
  maxRetainedBytes = 100 * 1024 * 1024, allowLoopbackForTest = false }: {
    host: string; port?: number; hostname: string; ttlMs?: number; maxRetainedBytes?: number; allowLoopbackForTest?: boolean }) {
  if (!Number.isFinite(maxRetainedBytes) || maxRetainedBytes < 0) throw new Error("Invalid retained attachment byte limit.");
  const addresses = networkInterfaces().tailscale0 ?? [];
  if (!(allowLoopbackForTest && host === "127.0.0.1") &&
      !addresses.some(address => address.address === host && address.family === "IPv4")) {
    throw new Error("Attachment server must bind to an IPv4 address on tailscale0.");
  }
  const dir = await mkdtemp(join(tmpdir(), "outlook-attachments-"));
  const files = new Map<string, { path: string; name: string; mime: string; size: number; expires: number; served: boolean; readers: number }>();
  let retainedBytes = 0;
  let queue = Promise.resolve();
  function serialize<T>(work: () => Promise<T>) {
    const result = queue.then(work);
    queue = result.then(() => {}, () => {});
    return result;
  }
  async function evict(needed = 0) {
    for (const expiredOnly of [true, false]) {
      for (const [key, file] of files) {
        if (file.readers || (expiredOnly ? file.expires > Date.now() : !file.served || retainedBytes + needed <= maxRetainedBytes)) continue;
        files.delete(key);
        try { await rm(file.path, { force: true }); } catch (error) { files.set(key, file); throw error; }
        retainedBytes -= file.size;
      }
    }
    if (retainedBytes + needed > maxRetainedBytes) throw new MailError("STORAGE_FULL", "Retained attachments fill the storage budget. Retry after links expire.");
  }
  const cleanup = () => serialize(() => evict());
  const server = createServer(async (req, res) => {
    res.setHeader("Cache-Control", "no-store");
    res.setHeader("X-Content-Type-Options", "nosniff");
    if (req.method !== "GET") return void res.writeHead(405, { Allow: "GET" }).end();
    const file = files.get(req.url ?? "");
    if (!file || file.expires <= Date.now()) return void res.writeHead(404).end();
    file.readers++;
    let released = false;
    const release = () => { if (!released) { released = true; file.readers--; } };
    const served = () => { file.served = true; release(); };
    res.once("finish", served);
    res.writeHead(200, { "Content-Type": file.mime, "Content-Length": file.size,
      "Content-Disposition": "attachment; filename*=UTF-8''" + encodeURIComponent(file.name) });
    try {
      await pipeline(createReadStream(file.path), res);
    } catch { res.destroy(); }
    finally { res.off("finish", served); release(); }
  });
  try {
    await new Promise<void>((resolve, reject) => server.once("error", reject).listen(port, host, resolve));
  } catch (error) {
    await rm(dir, { recursive: true, force: true });
    throw error;
  }
  const timer = setInterval(() => void cleanup().catch(error => console.error("Attachment cleanup failed", error)), 30_000);
  timer.unref();
  function save(name: string, size: number, mime: string, write: (path: string) => Promise<void>) {
    return serialize(async () => {
      if (size > maxRetainedBytes) throw new MailError("ATTACHMENT_TOO_LARGE", "Attachment exceeds the retained storage byte limit.");
      await evict(size);
      const token = randomBytes(32).toString("hex");
      const safeName = name.replace(/[\\/\x00-\x1f\x7f]/g, "_").slice(0, 200) || "attachment";
      const key = "/" + token + "/" + encodeURIComponent(safeName);
      const path = join(dir, token);
      try { await write(path); } catch (error) { await rm(path, { force: true }); throw error; }
      const expires = Date.now() + ttlMs;
      files.set(key, { path, name: safeName, mime, size, expires, served: false, readers: 0 });
      retainedBytes += size;
      const address = server.address();
      return { download_url: `http://${hostname}:${typeof address === "object" && address ? address.port : port}${key}`,
        expires_at: new Date(expires).toISOString() };
    });
  }
  return {
    async save(name: string, data: Buffer, mime: string) {
      return save(name, data.length, mime, path => writeFile(path, data, { mode: 0o600, flag: "wx" }));
    },
    async saveFile(name: string, source: string, mime: string) {
      const size = (await stat(source)).size;
      return save(name, size, mime, async path => {
        try { await rename(source, path); } catch (error) {
          if ((error as NodeJS.ErrnoException).code !== "EXDEV") throw error;
          await copyFile(source, path);
          await rm(source);
        }
        await chmod(path, 0o600);
      });
    },
    cleanup,
    directory: dir,
    async close() {
      clearInterval(timer);
      server.closeAllConnections();
      await new Promise<void>(resolve => server.close(() => resolve()));
      await queue;
      await rm(dir, { recursive: true, force: true });
    },
  };
}
