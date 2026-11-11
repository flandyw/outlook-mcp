import { execFile } from "node:child_process";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { extname, join } from "node:path";

export type AttachmentContent =
  | { kind: "text"; mime: string; format: string; text: string; truncated: boolean }
  | { kind: "image"; mime: string; base64: string }
  | { kind: "binary"; mime: string; note: string; base64?: string };

const imageTypes: [string, (data: Buffer) => boolean][] = [
  ["image/png", data => data.subarray(0, 8).equals(Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]))],
  ["image/jpeg", data => data[0] === 0xff && data[1] === 0xd8 && data[2] === 0xff],
  ["image/gif", data => /^GIF8[79]a/.test(data.subarray(0, 6).toString("latin1"))],
  ["image/webp", data => data.subarray(0, 4).toString("latin1") === "RIFF" && data.subarray(8, 12).toString("latin1") === "WEBP"],
];

const textExtensions = new Set([".txt", ".csv", ".tsv", ".md", ".json", ".xml", ".html", ".htm", ".ics", ".vcf",
  ".log", ".yaml", ".yml", ".ini", ".cfg", ".conf", ".eml", ".rtf", ".srt", ".vtt", ".sql", ".py", ".js", ".ts", ".css"]);

const mimeByExtension: Record<string, string> = {
  ".pdf": "application/pdf", ".txt": "text/plain", ".csv": "text/csv", ".html": "text/html", ".htm": "text/html",
  ".json": "application/json", ".xml": "application/xml", ".ics": "text/calendar", ".md": "text/markdown",
  ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  ".doc": "application/msword", ".xls": "application/vnd.ms-excel", ".ppt": "application/vnd.ms-powerpoint",
  ".zip": "application/zip", ".mp3": "audio/mpeg", ".mp4": "video/mp4", ".png": "image/png", ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg", ".gif": "image/gif", ".webp": "image/webp",
};

class ToolMissing extends Error {}

/** Runs a local converter with a time and output limit. Inputs never reach a shell. */
function run(command: string, args: string[]) {
  return new Promise<string>((resolve, reject) => {
    execFile(command, args, { timeout: 20_000, maxBuffer: 32 * 1024 * 1024, encoding: "utf8" }, (error, stdout) => {
      if ((error as NodeJS.ErrnoException | null)?.code === "ENOENT") reject(new ToolMissing(command));
      else if (error) reject(error);
      else resolve(stdout);
    });
  });
}

const entities: Record<string, string> = { amp: "&", lt: "<", gt: ">", quot: '"', apos: "'", nbsp: " " };
function decodeEntities(text: string) {
  return text.replace(/&(#x[0-9a-f]+|#\d+|[a-z]+);/gi, (whole, entity: string) => {
    if (entity[0] === "#") {
      const code = entity[1].toLowerCase() === "x" ? parseInt(entity.slice(2), 16) : parseInt(entity.slice(1), 10);
      return Number.isFinite(code) && code <= 0x10ffff ? String.fromCodePoint(code) : whole;
    }
    return entities[entity.toLowerCase()] ?? whole;
  });
}

const stripTags = (xml: string) => decodeEntities(xml.replace(/<[^>]+>/g, ""));
const tidy = (text: string) => text.replace(/[ \t]+\n/g, "\n").replace(/\n{3,}/g, "\n\n").trim();

function htmlToText(html: string) {
  return tidy(stripTags(html
    .replace(/<(script|style|head)\b[\s\S]*?<\/\1>/gi, "")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/(p|div|tr|li|h[1-6]|table|section|article)>/gi, "\n")
    .replace(/<\/t[dh]>/gi, "\t")));
}

async function withTempFile<T>(data: Buffer, name: string, work: (path: string) => Promise<T>) {
  const dir = await mkdtemp(join(tmpdir(), "outlook-mcp-attachment-"));
  try {
    const path = join(dir, "file" + extname(name).replace(/[^.a-z0-9]/gi, "").slice(0, 10));
    await writeFile(path, data, { mode: 0o600 });
    return await work(path);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
}

const zipEntries = async (path: string) => (await run("unzip", ["-Z1", path])).split("\n").filter(Boolean);
const zipRead = (path: string, entry: string) => run("unzip", ["-p", path, entry]);
const numbered = (entries: string[], pattern: RegExp) => entries
  .map(entry => ({ entry, n: Number(pattern.exec(entry)?.[1]) }))
  .filter(item => Number.isFinite(item.n)).sort((a, b) => a.n - b.n);

async function docxText(path: string) {
  const xml = await zipRead(path, "word/document.xml");
  return tidy(stripTags(xml.replace(/<w:tab\/>/g, "\t").replace(/<w:br[^>]*\/>/g, "\n").replace(/<\/w:p>/g, "\n")));
}

async function pptxText(path: string) {
  const slides = numbered(await zipEntries(path), /^ppt\/slides\/slide(\d+)\.xml$/);
  const parts: string[] = [];
  for (const { entry, n } of slides) {
    const text = tidy(stripTags((await zipRead(path, entry)).replace(/<a:br\/>/g, "\n").replace(/<\/a:p>/g, "\n")));
    parts.push("--- Slide " + n + " ---\n" + text);
  }
  return parts.join("\n\n");
}

function columnIndex(reference: string) {
  let index = 0;
  for (const letter of reference.replace(/\d+$/, "").toUpperCase()) index = index * 26 + letter.charCodeAt(0) - 64;
  return index - 1;
}

async function xlsxText(path: string) {
  const entries = await zipEntries(path);
  const shared: string[] = [];
  if (entries.includes("xl/sharedStrings.xml")) {
    for (const [, item] of (await zipRead(path, "xl/sharedStrings.xml")).matchAll(/<si>([\s\S]*?)<\/si>/g)) {
      shared.push(stripTags(item.replace(/<rPh\b[\s\S]*?<\/rPh>/g, "")));
    }
  }
  const names = entries.includes("xl/workbook.xml")
    ? [...(await zipRead(path, "xl/workbook.xml")).matchAll(/<sheet\b[^>]*\bname="([^"]*)"/g)].map(match => decodeEntities(match[1]))
    : [];
  const sheets = numbered(entries, /^xl\/worksheets\/sheet(\d+)\.xml$/);
  const parts: string[] = [];
  for (const [position, { entry, n }] of sheets.entries()) {
    const lines: string[] = [];
    for (const [, row] of (await zipRead(path, entry)).matchAll(/<row\b[^>]*>([\s\S]*?)<\/row>/g)) {
      const cells: string[] = [];
      for (const [, attributes, inner] of row.matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
        const reference = /\br="([A-Z]+\d+)"/.exec(attributes)?.[1];
        const type = /\bt="(\w+)"/.exec(attributes)?.[1];
        const raw = /<v>([\s\S]*?)<\/v>/.exec(inner ?? "")?.[1];
        let value = type === "inlineStr" ? stripTags(inner ?? "") : raw === undefined ? "" : decodeEntities(raw);
        if (type === "s") value = shared[Number(value)] ?? "";
        if (type === "b") value = value === "1" ? "TRUE" : "FALSE";
        const column = reference ? columnIndex(reference) : cells.length;
        while (cells.length < column) cells.push("");
        cells[column] = value.replace(/[\t\n]+/g, " ");
      }
      if (cells.some(Boolean)) lines.push(cells.join("\t").replace(/\t+$/, ""));
    }
    const name = names.length === sheets.length ? names[position] : "Sheet " + n;
    parts.push("--- " + name + " ---\n" + lines.join("\n"));
  }
  return parts.join("\n\n");
}

function looksLikeText(data: Buffer) {
  const sample = data.subarray(0, 8192);
  if (sample.includes(0)) return false;
  return !new TextDecoder("utf-8", { fatal: false }).decode(sample).includes("�");
}

/**
 * Turns attachment bytes into something a model can use: extracted text for documents,
 * the image itself for pictures, and optionally base64 for anything else.
 */
export async function extractAttachment(name: string, data: Buffer, options: { maxChars: number; includeData: boolean; maxInlineBytes: number }) {
  const extension = extname(name).toLowerCase();
  const fallbackMime = mimeByExtension[extension] ?? "application/octet-stream";
  const text = (format: string, mime: string, value: string): AttachmentContent => {
    const chars = Array.from(value);
    return { kind: "text", mime, format, text: chars.slice(0, options.maxChars).join(""), truncated: chars.length > options.maxChars };
  };
  const binary = (note: string, mime = fallbackMime): AttachmentContent => options.includeData && data.length <= options.maxInlineBytes
    ? { kind: "binary", mime, note, base64: data.toString("base64") }
    : { kind: "binary", mime, note: note + (options.includeData ? " The file is too large to return inline." : " Pass include_data: true for a base64 copy.") };

  const image = imageTypes.find(([, test]) => test(data));
  if (image) {
    return data.length <= options.maxInlineBytes ? { kind: "image", mime: image[0], base64: data.toString("base64") } as AttachmentContent
      : binary("Image is too large to return inline.", image[0]);
  }
  try {
    if (data.subarray(0, 5).toString("latin1") === "%PDF-") {
      const value = await withTempFile(data, name, path => run("pdftotext", ["-layout", "-enc", "UTF-8", "-q", path, "-"]));
      return value.trim() ? text("pdf", "application/pdf", tidy(value))
        : binary("The PDF has no text layer (it may be scanned).", "application/pdf");
    }
    if (data.subarray(0, 4).equals(Buffer.from([0x50, 0x4b, 0x03, 0x04]))) {
      const office = await withTempFile(data, name, async path => {
        const entries = await zipEntries(path);
        if (entries.includes("word/document.xml")) return ["docx", await docxText(path)];
        if (entries.some(entry => entry.startsWith("ppt/slides/"))) return ["pptx", await pptxText(path)];
        if (entries.includes("xl/workbook.xml")) return ["xlsx", await xlsxText(path)];
        return ["zip", "Archive contents:\n" + entries.slice(0, 500).join("\n")];
      });
      const mime = office[0] === "zip" ? "application/zip" : mimeByExtension["." + office[0]];
      return text(office[0], mime, office[1]);
    }
  } catch (error) {
    if (error instanceof ToolMissing) {
      return binary("Text extraction needs " + (error.message === "pdftotext" ? "pdftotext (poppler)" : "unzip") + " on the server host.");
    }
    return binary("Could not extract text from this file.");
  }
  if (textExtensions.has(extension) || looksLikeText(data)) {
    const decoded = data.toString("utf8").replace(/^﻿/, "");
    if (extension === ".html" || extension === ".htm" || /^\s*<(!doctype html|html)/i.test(decoded)) {
      return text("html", "text/html", htmlToText(decoded));
    }
    return text("text", fallbackMime === "application/octet-stream" ? "text/plain" : fallbackMime, decoded);
  }
  return binary("No text extractor for this file type.");
}
