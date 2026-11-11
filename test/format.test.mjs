import { test } from "node:test";
import assert from "node:assert/strict";
import { formatBody } from "../dist/format.js";

test("formats parts as Markdown and neutralises active syntax", () => {
  const { body_markdown, truncated } = formatBody([
    { kind: "paragraph", text: "  # Not a heading  " },
    { kind: "paragraph", text: "Mail first_last@example.com! [click](https://evil.test) <script>" },
    { kind: "bullet", text: "one\ntwo" },
    { kind: "quote", text: "quoted\nreply" },
    { kind: "separator", text: "" },
    { kind: "separator", text: "" },
    { kind: "paragraph", text: "---" },
    { kind: "separator", text: "" },
  ]);
  assert.equal(truncated, false);
  assert.equal(body_markdown, [
    "\\# Not a heading",
    "Mail first_last@example.com! \\[click\\](https://evil.test) \\<script\\>",
    "- one\n  two",
    "> quoted\n> reply",
    "---",
    "\\---",
  ].join("\n\n"));
});

test("truncates on code points", () => {
  const result = formatBody([{ kind: "paragraph", text: "😀".repeat(20) }], 5);
  assert.deepEqual(result, { body_markdown: "😀".repeat(5), truncated: true });
});
