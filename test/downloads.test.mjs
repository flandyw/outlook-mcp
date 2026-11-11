import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readdir } from 'node:fs/promises';
import { startDownloads } from '../dist/downloads.js';

test('download links preserve bytes, reject other paths and methods, expire and clean up', async () => {
  const server = await startDownloads({ host: '127.0.0.1', hostname: '127.0.0.1', port: 0, ttlMs: 150, allowLoopbackForTest: true });
  try {
    const data = Buffer.from('%PDF-1.7\nhello');
    const link = await server.save('../../Cover (2026).pdf', data, 'application/pdf');
    assert.match(new URL(link.download_url).pathname, /^\/[a-f0-9]{64}\//);
    const response = await fetch(link.download_url);
    assert.equal(response.status, 200);
    assert.equal(response.headers.get('content-type'), 'application/pdf');
    assert.deepEqual(Buffer.from(await response.arrayBuffer()), data);
    assert.equal((await fetch(link.download_url, { method: 'POST' })).status, 405);
    const origin = new URL(link.download_url).origin;
    for (const path of ['/', '/other', '/..%2f..%2fetc%2fpasswd']) assert.equal((await fetch(origin + path)).status, 404);
    await new Promise(resolve => setTimeout(resolve, 170));
    assert.equal((await fetch(link.download_url)).status, 404);
    await server.cleanup();
    assert.deepEqual(await readdir(server.directory), []);
  } finally { await server.close(); }
});

test('refuses non-Tailscale bind addresses', async () => {
  for (const host of ['0.0.0.0', '127.0.0.1', '192.168.1.82']) {
    await assert.rejects(startDownloads({host, hostname: host}), /tailscale0/);
  }
});
