import { test } from "node:test";
import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";

test("rejects unknown and inherited command names with usage", () => {
  for (const command of ["unknown", "constructor", "toString", "__proto__"]) {
    const result = spawnSync(process.execPath, ["dist/cli.js", command], { encoding: "utf8" });
    assert.equal(result.status, 1);
    assert.match(result.stderr, /^Usage:/);
    assert.equal(result.stdout, "");
  }
});
