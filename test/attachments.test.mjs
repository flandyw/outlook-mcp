import { test } from "node:test";
import assert from "node:assert/strict";
import { packageAttachment } from "../dist/attachments.js";

const png = Buffer.concat([Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), Buffer.alloc(16)]);
const jpeg = Buffer.from([0xff, 0xd8, 0xff, 0xe0, 1, 2, 3]);
const pdf = Buffer.from("%PDF-1.4\n%\xe2\xe3\xcf\xd3\n1 0 obj\n", "latin1");

test("returns images as images, sniffing the type from the bytes", () => {
  assert.deepEqual(packageAttachment("chart.png", png), { kind: "image", mime: "image/png", base64: png.toString("base64") });
  // Outlook can name an image without an extension.
  assert.deepEqual(packageAttachment("IMG_1234", jpeg), { kind: "image", mime: "image/jpeg", base64: jpeg.toString("base64") });
});

test("returns documents and other binaries unchanged as base64 files", () => {
  assert.deepEqual(packageAttachment("report.pdf", pdf), { kind: "file", mime: "application/pdf", base64: pdf.toString("base64") });
  assert.deepEqual(packageAttachment("misnamed.txt", pdf).mime, "application/pdf", "content wins over a wrong extension");
  const docx = Buffer.from([0x50, 0x4b, 0x03, 0x04, 0, 0]);
  assert.deepEqual(packageAttachment("Letter.DOCX", docx), {
    kind: "file", mime: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", base64: docx.toString("base64"),
  });
  const unknown = Buffer.from([0, 1, 2, 255]);
  assert.deepEqual(packageAttachment("blob.bin", unknown), { kind: "file", mime: "application/octet-stream", base64: unknown.toString("base64") });
});

test("returns UTF-8 text files as text, falling back to base64 when they are not valid UTF-8", () => {
  assert.deepEqual(packageAttachment("data.csv", Buffer.from("﻿a,b\n1,2\n")), { kind: "text", mime: "text/csv", text: "a,b\n1,2\n" });
  assert.deepEqual(packageAttachment("invite.ics", Buffer.from("BEGIN:VCALENDAR")), { kind: "text", mime: "text/calendar", text: "BEGIN:VCALENDAR" });
  assert.deepEqual(packageAttachment("notes", Buffer.from("plain words")), { kind: "text", mime: "text/plain", text: "plain words" });
  const latin1 = Buffer.from("caf\xe9", "latin1");
  assert.deepEqual(packageAttachment("old.txt", latin1), { kind: "file", mime: "text/plain", base64: latin1.toString("base64") });
});
