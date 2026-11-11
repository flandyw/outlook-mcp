import { test } from "node:test";
import assert from "node:assert/strict";
import { crc32 } from "node:zlib";
import { extractAttachment } from "../dist/attachments.js";

const options = { maxChars: 50_000, includeData: false, maxInlineBytes: 1024 * 1024 };

// Minimal uncompressed ZIP writer, enough for Office Open XML fixtures.
function zip(files) {
  const locals = [], centrals = [];
  let offset = 0;
  for (const [name, text] of Object.entries(files)) {
    const nameBytes = Buffer.from(name), data = Buffer.from(text), crc = crc32(data);
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0); local.writeUInt16LE(20, 4); local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(data.length, 18); local.writeUInt32LE(data.length, 22); local.writeUInt16LE(nameBytes.length, 26);
    const central = Buffer.alloc(46);
    central.writeUInt32LE(0x02014b50, 0); central.writeUInt16LE(20, 4); central.writeUInt16LE(20, 6); central.writeUInt32LE(crc, 16);
    central.writeUInt32LE(data.length, 20); central.writeUInt32LE(data.length, 24); central.writeUInt16LE(nameBytes.length, 28);
    central.writeUInt32LE(offset, 42);
    locals.push(local, nameBytes, data);
    centrals.push(central, nameBytes);
    offset += local.length + nameBytes.length + data.length;
  }
  const directory = Buffer.concat(centrals), end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0); end.writeUInt16LE(Object.keys(files).length, 8); end.writeUInt16LE(Object.keys(files).length, 10);
  end.writeUInt32LE(directory.length, 12); end.writeUInt32LE(offset, 16);
  return Buffer.concat([...locals, directory, end]);
}

function pdf(text) {
  const stream = `BT /F1 12 Tf 72 720 Td (${text}) Tj ET`;
  const objects = ["<< /Type /Catalog /Pages 2 0 R >>", "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>",
    `<< /Length ${stream.length} >>\nstream\n${stream}\nendstream`, "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>"];
  let body = "%PDF-1.4\n";
  const offsets = objects.map((object, index) => { const at = body.length; body += `${index + 1} 0 obj\n${object}\nendobj\n`; return at; });
  const xref = body.length;
  body += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n` + offsets.map(at => String(at).padStart(10, "0") + " 00000 n \n").join("");
  body += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  return Buffer.from(body, "latin1");
}

test("extracts text from PDF, Word, PowerPoint and Excel files", async () => {
  const fromPdf = await extractAttachment("report.pdf", pdf("Quarterly totals"), options);
  assert.equal(fromPdf.kind, "text");
  assert.equal(fromPdf.format, "pdf");
  assert.match(fromPdf.text, /Quarterly totals/);

  const docx = await extractAttachment("letter.docx", zip({ "word/document.xml":
    '<w:document><w:body><w:p><w:r><w:t>Dear Sam,</w:t></w:r></w:p><w:p><w:r><w:t>Fish &amp; chips</w:t><w:tab/><w:t>today</w:t></w:r></w:p></w:body></w:document>' }), options);
  assert.equal(docx.format, "docx");
  assert.equal(docx.text, "Dear Sam,\nFish & chips\ttoday");

  const pptx = await extractAttachment("deck.pptx", zip({
    "ppt/slides/slide2.xml": "<p:sld><a:p><a:r><a:t>Second</a:t></a:r></a:p></p:sld>",
    "ppt/slides/slide10.xml": "<p:sld><a:p><a:r><a:t>Tenth</a:t></a:r></a:p></p:sld>",
    "ppt/slides/slide1.xml": "<p:sld><a:p><a:r><a:t>First</a:t></a:r></a:p></p:sld>",
  }), options);
  assert.equal(pptx.text, "--- Slide 1 ---\nFirst\n\n--- Slide 2 ---\nSecond\n\n--- Slide 10 ---\nTenth");

  const xlsx = await extractAttachment("budget.xlsx", zip({
    "xl/workbook.xml": '<workbook><sheets><sheet name="Budget" sheetId="1" r:id="rId1"/></sheets></workbook>',
    "xl/sharedStrings.xml": "<sst><si><t>Item</t></si><si><t>Cost</t></si><si><r><t>Ca</t></r><r><t>ke</t></r></si></sst>",
    "xl/worksheets/sheet1.xml": '<worksheet><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="C1" t="s"><v>1</v></c></row>' +
      '<row r="2"><c r="A2" t="s"><v>2</v></c><c r="C2"><v>12.5</v></c><c r="D2" t="b"><v>1</v></c></row></sheetData></worksheet>',
  }), options);
  assert.equal(xlsx.text, "--- Budget ---\nItem\t\tCost\nCake\t\t12.5\tTRUE");
});

test("handles text, HTML, images, other files and limits", async () => {
  const csv = await extractAttachment("data.csv", Buffer.from("﻿a,b\n1,2\n"), options);
  assert.deepEqual([csv.kind, csv.mime, csv.text], ["text", "text/csv", "a,b\n1,2\n"]);
  const html = await extractAttachment("page.html", Buffer.from("<html><style>x{}</style><p>Hi &amp; bye</p><table><tr><td>a</td><td>b</td></tr></table></html>"), options);
  assert.equal(html.text, "Hi & bye\na\tb");
  const png = Buffer.concat([Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), Buffer.alloc(16)]);
  const image = await extractAttachment("photo.png", png, options);
  assert.deepEqual([image.kind, image.mime, image.base64], ["image", "image/png", png.toString("base64")]);
  const bytes = Buffer.from([0, 1, 2, 3, 255]);
  const plain = await extractAttachment("song.mp3", bytes, options);
  assert.deepEqual([plain.kind, plain.mime, plain.base64], ["binary", "audio/mpeg", undefined]);
  assert.match(plain.note, /include_data/);
  const withData = await extractAttachment("song.mp3", bytes, { ...options, includeData: true });
  assert.equal(withData.base64, bytes.toString("base64"));
  const capped = await extractAttachment("long.txt", Buffer.from("é".repeat(30)), { ...options, maxChars: 10 });
  assert.deepEqual([capped.text, capped.truncated], ["é".repeat(10), true]);
  const scanned = await extractAttachment("scan.pdf", pdf(""), options);
  assert.equal(scanned.kind, "binary");
  assert.match(scanned.note, /no text layer/);
});

test("explains a missing converter instead of failing", async t => {
  const path = process.env.PATH;
  t.after(() => { process.env.PATH = path; });
  process.env.PATH = "/nonexistent";
  const result = await extractAttachment("report.pdf", pdf("x"), options);
  assert.equal(result.kind, "binary");
  assert.match(result.note, /pdftotext/);
});
