export interface BodyPart {
  kind: "text" | "paragraph" | "bullet" | "quote" | "row" | "separator";
  text: string;
}

// Neutralise Markdown that could render as links, HTML, code or headings, while keeping
// ordinary punctuation (underscores in addresses, asterisks, exclamation marks) readable.
const escape = (value: string) => value
  .replace(/[\\`\[\]<>]/g, "\\$&")
  .replace(/^(\s*)(#{1,6}|[-+*]|\d+[.)])(?=\s)/gm, "$1\\$2")
  .replace(/^(\s*)(={3,}|-{3,}|_{3,}|\*{3,})\s*$/gm, "$1\\$2");

const clean = (text: string) => text
  .split("\n").map(line => line.replace(/[ \t ]+/g, " ").trim()).join("\n")
  .replace(/\n{3,}/g, "\n\n").trim();

export function formatBody(parts: BodyPart[], limit = 12000) {
  const blocks: string[] = [];
  for (const { kind, text } of parts) {
    if (kind === "separator") {
      if (blocks.length && blocks[blocks.length - 1] !== "---") blocks.push("---");
      continue;
    }
    const safe = escape(clean(text));
    if (!safe) continue;
    if (kind === "bullet") blocks.push("- " + safe.replace(/\n/g, "\n  "));
    else if (kind === "quote") blocks.push(safe.split("\n").map(line => "> " + line).join("\n"));
    else blocks.push(safe);
  }
  while (blocks[blocks.length - 1] === "---") blocks.pop();
  const markdown = blocks.join("\n\n");
  const chars = Array.from(markdown);
  if (chars.length <= limit) return { body_markdown: markdown, truncated: false };
  return { body_markdown: chars.slice(0, limit).join("").trimEnd(), truncated: true };
}
