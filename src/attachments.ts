import { extname } from "node:path";

/**
 * Attachments are returned as the file itself; the model or client interprets them.
 * This module only works out a MIME type and whether the bytes are plain UTF-8 text.
 */
export type AttachmentPayload =
  | { kind: "image"; mime: string; base64: string }
  | { kind: "text"; mime: string; text: string }
  | { kind: "file"; mime: string; base64: string };

const signatures: [string, (data: Buffer) => boolean][] = [
  ["image/png", data => data.subarray(0, 8).equals(Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]))],
  ["image/jpeg", data => data[0] === 0xff && data[1] === 0xd8 && data[2] === 0xff],
  ["image/gif", data => /^GIF8[79]a/.test(data.subarray(0, 6).toString("latin1"))],
  ["image/webp", data => data.subarray(0, 4).toString("latin1") === "RIFF" && data.subarray(8, 12).toString("latin1") === "WEBP"],
  ["application/pdf", data => data.subarray(0, 5).toString("latin1") === "%PDF-"],
];

const byExtension: Record<string, string> = {
  ".pdf": "application/pdf", ".txt": "text/plain", ".csv": "text/csv", ".tsv": "text/tab-separated-values",
  ".html": "text/html", ".htm": "text/html", ".md": "text/markdown", ".json": "application/json", ".xml": "application/xml",
  ".ics": "text/calendar", ".vcf": "text/vcard", ".eml": "message/rfc822", ".rtf": "application/rtf",
  ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  ".doc": "application/msword", ".xls": "application/vnd.ms-excel", ".ppt": "application/vnd.ms-powerpoint",
  ".zip": "application/zip", ".mp3": "audio/mpeg", ".m4a": "audio/mp4", ".wav": "audio/wav", ".mp4": "video/mp4",
  ".mov": "video/quicktime", ".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".gif": "image/gif",
  ".webp": "image/webp", ".heic": "image/heic", ".svg": "image/svg+xml",
};

// Text-like types that are sent as readable text rather than base64.
const textMime = (mime: string) => mime.startsWith("text/") ||
  ["application/json", "application/xml", "message/rfc822", "image/svg+xml"].includes(mime);

function isUtf8Text(data: Buffer) {
  if (data.subarray(0, 8192).includes(0)) return false;
  try {
    new TextDecoder("utf-8", { fatal: true }).decode(data);
    return true;
  } catch {
    return false;
  }
}

export function packageAttachment(name: string, data: Buffer): AttachmentPayload {
  const sniffed = signatures.find(([, test]) => test(data))?.[0];
  const mime = sniffed ?? byExtension[extname(name).toLowerCase()] ?? (isUtf8Text(data) ? "text/plain" : "application/octet-stream");
  if (mime.startsWith("image/") && mime !== "image/svg+xml" && mime !== "image/heic") {
    return { kind: "image", mime, base64: data.toString("base64") };
  }
  if (textMime(mime) && isUtf8Text(data)) return { kind: "text", mime, text: data.toString("utf8").replace(/^﻿/, "") };
  return { kind: "file", mime, base64: data.toString("base64") };
}
