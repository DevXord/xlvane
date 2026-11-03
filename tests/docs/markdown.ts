// Reads fenced code blocks from the Markdown guides that tests/docs checks.

/** The fenced blocks of a Markdown text whose info string matches `lang`, with the 1-based line of their first line of code. */
export function codeBlocks(text: string, lang: RegExp): { line: number; code: string }[] {
  const lines = text.split('\n');
  const out: { line: number; code: string }[] = [];
  for (let i = 0; i < lines.length; i++) {
    const m = /^```(\S*)\s*$/.exec(lines[i]!);
    if (!m) continue;
    const start = i + 1;
    while (i + 1 < lines.length && lines[i + 1] !== '```') i++;
    if (lang.test(m[1]!)) out.push({ line: start + 1, code: lines.slice(start, i + 1).join('\n') });
    i++;
  }
  return out;
}
