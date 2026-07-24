/**
 * Markdown to Teams HTML conversion utilities.
 */

import { escapeHtmlChars } from './parsers-html.js';

/**
 * Converts inline markdown formatting to Teams HTML within a single line.
 * Handles: bold, italic, strikethrough, inline code.
 * Text outside of formatting markers is HTML-escaped.
 */
function convertInlineFormatting(line: string): string {
  // Process inline code first (to prevent other formatting inside code spans)
  // Split on `code` patterns, escape and format alternately
  const codeParts = line.split(/`([^`]+)`/);
  let result = '';
  
  for (let i = 0; i < codeParts.length; i++) {
    if (i % 2 === 1) {
      // Inside backticks - render as code, only escape HTML
      result += `<code>${escapeHtmlChars(codeParts[i])}</code>`;
    } else {
      // Outside backticks - process other inline formatting
      let segment = escapeHtmlChars(codeParts[i]);
      
      // Bold: **text** or __text__
      segment = segment.replace(/\*\*(.+?)\*\*/g, '<b>$1</b>');
      segment = segment.replace(/__(.+?)__/g, '<b>$1</b>');
      
      // Italic: *text* or _text_ (but not inside words for underscore)
      segment = segment.replace(/\*(.+?)\*/g, '<i>$1</i>');
      segment = segment.replace(/(?<!\w)_(.+?)_(?!\w)/g, '<i>$1</i>');
      
      // Strikethrough: ~~text~~
      segment = segment.replace(/~~(.+?)~~/g, '<s>$1</s>');
      
      result += segment;
    }
  }
  
  return result;
}

/**
 * Matches a markdown table separator row, e.g. `| --- | :---: |`.
 */
function isTableSeparator(line: string): boolean {
  return /^\s*\|?\s*:?-+:?\s*(\|\s*:?-+:?\s*)*\|?\s*$/.test(line);
}

/**
 * Converts a markdown pipe table (header row, a separator row of dashes, then
 * body rows) into a Teams HTML <table>.
 */
function buildTableHtml(lines: string[]): string {
  const parseRow = (l: string): string[] =>
    l.replace(/^\s*\|/, '').replace(/\|\s*$/, '').split('|').map(c => c.trim());
  const header = parseRow(lines[0]);
  const bodyRows = lines.slice(2);
  const headHtml = `<thead><tr>${header.map(c => `<th>${convertInlineFormatting(c)}</th>`).join('')}</tr></thead>`;
  const bodyHtml = bodyRows
    .map(l => `<tr>${parseRow(l).map(c => `<td>${convertInlineFormatting(c)}</td>`).join('')}</tr>`)
    .join('');
  return `<table>${headHtml}<tbody>${bodyHtml}</tbody></table>`;
}

/**
 * Returns true when a raw markdown line (before HTML conversion) consists
 * solely of a bold span — e.g. "**Target**" or "__Label__" — with optional
 * leading whitespace (e.g. an indented label) and trailing whitespace / hard-break
 * markers ("  "). These label lines should be flushed as their own <p> rather
 * than joined to the next line with a <br>, so headings sit directly above
 * their content without cramping.
 */
function isBoldOnlyLine(raw: string): boolean {
  const trimmed = raw.trim();
  return /^(\*\*[^*]+\*\*|__[^_]+__)$/.test(trimmed);
}

/**
 * Renders the lines of a single block (text already split on blank lines) into
 * Teams HTML. Walks the lines and groups consecutive runs by type, so block
 * elements (lists, blockquotes, tables, headings) can interrupt a paragraph
 * without a separating blank line, matching standard markdown behaviour. Plain
 * lines accumulate into a paragraph and are joined with <br>.
 *
 * Trailing two-space hard-break markers ("  ") are stripped from every line up
 * front so all branches (lists, blockquotes, tables, plain text) see the clean
 * version — they are a Teams/LLM workaround that is no longer needed now that
 * bold-only lines are treated as their own block.
 *
 * A line that contains only a bold span — e.g. "**Target**" — is flushed as
 * its own <p> so the next line always starts a fresh paragraph, preventing the
 * heading and its content from being crammed together with a <br>.
 */
function renderTextBlock(rawLines: string[]): string {
  // Strip trailing hard-break markers ("  ") from every line up front so all
  // subsequent branches (headings, lists, blockquotes, tables, plain text) see
  // the clean version without needing to repeat the stripping themselves.
  const lines = rawLines.map(l => l.replace(/  +$/, ''));

  const out: string[] = [];
  let para: string[] = [];
  const flushParagraph = (): void => {
    if (para.length) {
      out.push(`<p>${para.join('<br>')}</p>`);
      para = [];
    }
  };

  let i = 0;
  while (i < lines.length) {
    const line = lines[i];

    // Heading: its own single-line block
    const heading = line.match(/^\s*(#{1,6})\s+(.*)$/);
    if (heading) {
      flushParagraph();
      out.push(`<h${heading[1].length}>${convertInlineFormatting(heading[2])}</h${heading[1].length}>`);
      i++;
      continue;
    }

    // Table: a row containing a pipe immediately followed by a separator row
    if (i + 1 < lines.length && /\|/.test(line) && isTableSeparator(lines[i + 1])) {
      flushParagraph();
      const tableLines = [line, lines[i + 1]];
      i += 2;
      while (i < lines.length && /\|/.test(lines[i]) && !isTableSeparator(lines[i])) {
        tableLines.push(lines[i]);
        i++;
      }
      out.push(buildTableHtml(tableLines));
      continue;
    }

    // Unordered list run
    if (/^\s*[-*]\s+/.test(line)) {
      flushParagraph();
      const items: string[] = [];
      while (i < lines.length && /^\s*[-*]\s+/.test(lines[i])) {
        items.push(`<li>${convertInlineFormatting(lines[i].replace(/^\s*[-*]\s+/, ''))}</li>`);
        i++;
      }
      out.push(`<ul>${items.join('')}</ul>`);
      continue;
    }

    // Ordered list run
    if (/^\s*\d+[.)]\s+/.test(line)) {
      flushParagraph();
      const items: string[] = [];
      while (i < lines.length && /^\s*\d+[.)]\s+/.test(lines[i])) {
        items.push(`<li>${convertInlineFormatting(lines[i].replace(/^\s*\d+[.)]\s+/, ''))}</li>`);
        i++;
      }
      out.push(`<ol>${items.join('')}</ol>`);
      continue;
    }

    // Blockquote run
    if (/^\s*>\s?/.test(line)) {
      flushParagraph();
      const inner: string[] = [];
      while (i < lines.length && /^\s*>\s?/.test(lines[i])) {
        inner.push(convertInlineFormatting(lines[i].replace(/^\s*>\s?/, '')));
        i++;
      }
      out.push(`<blockquote>${inner.join('<br>')}</blockquote>`);
      continue;
    }

    // Bold-only label line (e.g. "**Target**"): flush as its own <p> so the
    // next line starts a fresh paragraph rather than being joined with <br>.
    if (isBoldOnlyLine(line)) {
      flushParagraph();
      out.push(`<p>${convertInlineFormatting(line)}</p>`);
      i++;
      continue;
    }

    // Plain text line: accumulate into the current paragraph
    para.push(convertInlineFormatting(line));
    i++;
  }

  flushParagraph();
  return out.join('');
}

/**
 * Joins rendered top-level blocks into the final HTML.
 *
 * A blank line in the source (`\n\n`) is an intentional paragraph break. Teams'
 * `RichText/Html` chat renderer collapses the margin between adjacent `<p>`
 * elements, so `<p>A</p><p>B</p>` arrives with no visible gap (cramped). A
 * `<br><br>` inside a single `<p>`, on the other hand, DOES render a blank line.
 *
 * So when one block ends a paragraph (`</p>`) and the next begins one (`<p>`),
 * which only happens across a blank-line boundary (a single source block is
 * rendered as one entry), we merge them into one `<p>` separated by `<br><br>`
 * to produce the visible gap the author intended. Genuine block elements (lists,
 * tables, headings, code, blockquotes) keep their own boundaries untouched, and
 * a heading directly above its content (no blank line) stays tight because that
 * happens inside a single rendered block, not at a join seam.
 */
function joinBlocksWithParagraphGaps(parts: string[]): string {
  let html = '';
  for (const part of parts) {
    if (!part) continue;
    if (html.endsWith('</p>') && part.startsWith('<p>')) {
      // Drop the seam "</p><p>" and stitch the paragraphs together with a gap.
      html = `${html.slice(0, -4)}<br><br>${part.slice(3)}`;
    } else {
      html += part;
    }
  }
  return html;
}

/**
 * Converts markdown-formatted text to Teams-compatible HTML.
 * 
 * Supports:
 * - **bold** / __bold__ → <b>
 * - *italic* / _italic_ → <i>
 * - ~~strikethrough~~ → <s>
 * - `inline code` → <code>
 * - ```code blocks``` → <pre><code>
 * - Blank line (\n\n) → visible paragraph gap; single newline (\n) → line break
 * - Ordered lists (1. item) → <ol><li>
 * - Unordered lists (- item, * item) → <ul><li>
 * - Headings (#..###### ) → <h1>..<h6>
 * - Blockquotes (> line) → <blockquote>
 * - Pipe tables (| a | b | with a --- separator row) → <table>
 * 
 * Plain text without any formatting is returned as-is (HTML-escaped).
 */
export function markdownToTeamsHtml(text: string): string {
  // Handle fenced code blocks first (```...```)
  // Split text into code blocks and non-code-block segments
  const segments: { type: 'text' | 'codeblock'; content: string; lang?: string }[] = [];
  const codeBlockRegex = /```(\w*)\n?([\s\S]*?)```/g;
  let lastIndex = 0;
  let match;
  
  while ((match = codeBlockRegex.exec(text)) !== null) {
    // Text before this code block
    if (match.index > lastIndex) {
      segments.push({ type: 'text', content: text.substring(lastIndex, match.index) });
    }
    segments.push({ type: 'codeblock', content: match[2], lang: match[1] || undefined });
    lastIndex = match.index + match[0].length;
  }
  // Remaining text after last code block
  if (lastIndex < text.length) {
    segments.push({ type: 'text', content: text.substring(lastIndex) });
  }
  
  const htmlParts: string[] = [];
  
  for (const segment of segments) {
    if (segment.type === 'codeblock') {
      // Code blocks: escape HTML but preserve whitespace
      const escaped = escapeHtmlChars(segment.content.replace(/\n$/, ''));
      htmlParts.push(`<pre><code>${escaped}</code></pre>`);
      continue;
    }
    
    // Process text segments: split into paragraphs on double newlines, then
    // render each block, letting block elements interrupt a paragraph.
    const paragraphs = segment.content.split(/\n{2,}/);
    
    for (const para of paragraphs) {
      const trimmed = para.trim();
      if (!trimmed) continue;
      htmlParts.push(renderTextBlock(trimmed.split('\n')));
    }
  }
  
  return joinBlocksWithParagraphGaps(htmlParts) || '<p></p>';
}

/**
 * Checks whether text contains any markdown formatting that would
 * benefit from conversion to HTML.
 */
export function hasMarkdownFormatting(text: string): boolean {
  // Code blocks
  if (/```[\s\S]*```/.test(text)) return true;
  // Inline code
  if (/`[^`]+`/.test(text)) return true;
  // Bold
  if (/\*\*.+?\*\*/.test(text) || /__.+?__/.test(text)) return true;
  // Italic (single * or _)
  if (/(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)/.test(text)) return true;
  // Strikethrough
  if (/~~.+?~~/.test(text)) return true;
  // Lists
  if (/^\s*[-*]\s+/m.test(text)) return true;
  if (/^\s*\d+[.)]\s+/m.test(text)) return true;
  // Multiple newlines (paragraph breaks)
  if (/\n/.test(text)) return true;
  
  return false;
}
