import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readdir, mkdtemp, writeFile, stat, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { getEventListeners } from 'node:events';
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

test('moves files without buffering, preserves bytes and makes the retained copy private', async t => {
  const dir = await mkdtemp(join(tmpdir(), 'attachment-move-'));
  t.after(() => rm(dir, { recursive: true, force: true }));
  const server = await startDownloads({ host: '127.0.0.1', hostname: '127.0.0.1', port: 0, maxRetainedBytes: 16, allowLoopbackForTest: true });
  t.after(() => server.close());
  const source = join(dir, 'file');
  await writeFile(source, 'original bytes', { mode: 0o644 });
  const before = await stat(source);
  const link = await server.saveFile('file.txt', source, 'text/plain');
  await assert.rejects(stat(source), { code: 'ENOENT' });
  const [name] = await readdir(server.directory);
  const after = await stat(join(server.directory, name));
  assert.equal(after.ino, before.ino, 'same-filesystem transfer must rename, not copy');
  assert.equal(after.mode & 0o777, 0o600);
  assert.equal(await (await fetch(link.download_url)).text(), 'original bytes');
  await assert.rejects(server.saveFile('missing.txt', source, 'text/plain'), { code: 'ENOENT' });
  const empty = await server.save('empty.txt', Buffer.alloc(0), 'text/plain');
  assert.equal(await (await fetch(empty.download_url)).text(), '');
  const controller = new AbortController();
  await writeFile(source, 'ok');
  const retained = await server.saveFile('retained.txt', source, 'text/plain', controller.signal);
  server.commit(retained.download_url);
  assert.equal(getEventListeners(controller.signal, 'abort').length, 0);
  controller.abort();
  assert.equal(await (await fetch(retained.download_url)).text(), 'ok', 'committed links survive normal request closure');
  await writeFile(source, 'no');
  await assert.rejects(server.saveFile('cancelled.txt', source, 'text/plain', controller.signal), { code: 'CANCELLED' });
  assert.equal((await stat(source)).size, 2);
});

test('retained-byte quota rejects unserved files and evicts oldest served files under pressure', async t => {
  const server = await startDownloads({ host: '127.0.0.1', hostname: '127.0.0.1', port: 0, maxRetainedBytes: 12, allowLoopbackForTest: true });
  t.after(() => server.close());
  const first = await server.save('first', Buffer.from('first!'), 'text/plain');
  const second = await server.save('second', Buffer.from('second'), 'text/plain');
  await assert.rejects(server.save('blocked', Buffer.from('third!'), 'text/plain'), { code: 'STORAGE_FULL' });
  await assert.rejects(server.save('huge', Buffer.alloc(13), 'text/plain'), { code: 'ATTACHMENT_TOO_LARGE' });
  assert.equal(await (await fetch(second.download_url)).text(), 'second');
  assert.equal(await (await fetch(first.download_url)).text(), 'first!');
  const third = await server.save('third', Buffer.from('third!'), 'text/plain');
  assert.equal((await fetch(first.download_url)).status, 404, 'oldest creation wins, not last-served order');
  assert.equal(await (await fetch(second.download_url)).text(), 'second');
  assert.equal(await (await fetch(third.download_url)).text(), 'third!');
  const names = await readdir(server.directory);
  const sizes = await Promise.all(names.map(name => stat(join(server.directory, name))));
  assert.equal(sizes.reduce((total, info) => total + info.size, 0), 12);
});

test('concurrent saves cannot exceed the retained-byte budget', async t => {
  const server = await startDownloads({ host: '127.0.0.1', hostname: '127.0.0.1', port: 0, maxRetainedBytes: 4, allowLoopbackForTest: true });
  t.after(() => server.close());
  const results = await Promise.allSettled([server.save('a', Buffer.alloc(3), 'text/plain'), server.save('b', Buffer.alloc(3), 'text/plain')]);
  assert.equal(results[0].status, 'fulfilled');
  assert.equal(results[1].reason.code, 'STORAGE_FULL');
  assert.equal((await readdir(server.directory)).length, 1);
});
