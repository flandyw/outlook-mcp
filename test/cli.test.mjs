import { test } from "node:test";
import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { chmod, mkdtemp, readdir, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

test("rejects unknown and inherited command names with usage", () => {
  for (const command of ["unknown", "constructor", "toString", "__proto__"]) {
    const result = spawnSync(process.execPath, ["dist/cli.js", command], { encoding: "utf8" });
    assert.equal(result.status, 1);
    assert.match(result.stderr, /^Usage:/);
    assert.equal(result.stdout, "");
  }
});

function runStatus(args, env) {
  return spawnSync(process.execPath, ["dist/cli.js", "status", ...args], {
    encoding: "utf8",
    env: {
      PATH: process.env.PATH ?? "",
      HOME: process.env.HOME ?? "",
      TMPDIR: process.env.TMPDIR ?? "/tmp",
      LANG: process.env.LANG ?? "C.UTF-8",
      ...env,
    },
  });
}

test("status validates the configured password file without printing it", async t => {
  const dir = await mkdtemp(join(tmpdir(), "outlook-status-"));
  t.after(async () => {
    for (const name of await readdir(dir).catch(() => [])) await chmod(join(dir, name), 0o700).catch(() => {});
    await rm(dir, { recursive: true, force: true });
  });
  const base = {
    OUTLOOK_CONTROL_SOCKET: join(dir, "missing.sock"),
    OUTLOOK_PROFILE: join(dir, "profile"),
  };
  const canary = "status-canary-do-not-print";
  const quiet = runStatus([], base);
  assert.equal(quiet.status, 0, quiet.stderr);
  assert.match(quiet.stdout, /not running/);
  assert.match(quiet.stdout, /Password: not configured/);
  assert.equal(quiet.stderr, "");

  const ready = join(dir, "ready");
  await writeFile(ready, canary + "\n");
  await chmod(ready, 0o600);
  const valid = runStatus([], { ...base, OUTLOOK_PASSWORD_FILE: ready });
  assert.equal(valid.status, 0, valid.stderr);
  assert.match(valid.stdout, /Password: valid \(/);
  assert.equal(valid.stdout.includes(ready), true);
  assert.equal((valid.stdout + valid.stderr).includes(canary), false);

  const missing = join(dir, "alpha-token-value");
  await writeFile(join(dir, "alpha-token-value-bak"), canary + "\n");
  await chmod(join(dir, "alpha-token-value-bak"), 0o600);
  await writeFile(join(dir, "notes.txt"), canary + "\n");
  for (const args of [[], ["--check"]]) {
    const result = runStatus(args, { ...base, OUTLOOK_PASSWORD_FILE: missing });
    assert.equal(result.status, 1);
    assert.match(result.stdout, /Password: invalid/);
    assert.match(result.stderr, /Password file misconfigured/);
    assert.match(result.stderr, /is missing/);
    assert.equal(result.stderr.includes(join(dir, "alpha-token-value-bak")), true);
    assert.match(result.stderr, /not used automatically/);
    assert.equal(result.stderr.includes(join(dir, "notes.txt")), false);
    assert.equal((result.stdout + result.stderr).includes(canary), false);
  }

  const loose = join(dir, "loose");
  await writeFile(loose, canary + "\n");
  await chmod(loose, 0o644);
  const looseResult = runStatus([], { ...base, OUTLOOK_PASSWORD_FILE: loose });
  assert.equal(looseResult.status, 1);
  assert.match(looseResult.stderr, /loose permissions/);
  assert.match(looseResult.stderr, /0644/);
  assert.equal((looseResult.stdout + looseResult.stderr).includes(canary), false);

  const empty = join(dir, "empty");
  await writeFile(empty, "\n");
  await chmod(empty, 0o600);
  const emptyResult = runStatus([], { ...base, OUTLOOK_PASSWORD_FILE: empty });
  assert.equal(emptyResult.status, 1);
  assert.match(emptyResult.stderr, /is empty/);
  assert.equal((emptyResult.stdout + emptyResult.stderr).includes(canary), false);
});
