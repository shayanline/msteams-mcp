/**
 * Dedicated unit tests for HTML parsing/escaping utilities.
 *
 * parsers.test.ts exercises some of these via the re-export barrel, but the
 * `<item>` link branches and escapeHtmlChars need direct coverage.
 */

import { describe, it, expect } from 'vitest';
import { extractLinks, stripHtml, escapeHtmlChars } from './parsers-html.js';

describe('extractLinks (anchor tags)', () => {
  it('extracts a standard anchor link', () => {
    const html = 'go <a href="https://example.com">here</a> now';
    expect(extractLinks(html)).toEqual([{ url: 'https://example.com', text: 'here' }]);
  });

  it('strips nested HTML from anchor text', () => {
    const html = '<a href="https://e.com"><b>Bold</b> text</a>';
    expect(extractLinks(html)).toEqual([{ url: 'https://e.com', text: 'Bold text' }]);
  });

  it('falls back to the url when anchor text is empty', () => {
    const html = '<a href="https://only-url.com"></a>';
    expect(extractLinks(html)).toEqual([{ url: 'https://only-url.com', text: 'https://only-url.com' }]);
  });

  it('ignores javascript: anchors', () => {
    expect(extractLinks('<a href="javascript:alert(1)">x</a>')).toEqual([]);
  });

  it('returns empty array when no links present', () => {
    expect(extractLinks('plain text')).toEqual([]);
    expect(extractLinks('')).toEqual([]);
  });
});

describe('extractLinks (Teams <item> URIs)', () => {
  it('extracts an item with type before uri', () => {
    const html = '<item type="amsTranscript" uri="https://t.example/transcript">';
    expect(extractLinks(html)).toEqual([
      { url: 'https://t.example/transcript', text: 'amsTranscript', contentType: 'amsTranscript' },
    ]);
  });

  it('extracts an item with uri before type', () => {
    const html = '<item uri="https://t.example/video" type="onedriveForBusinessVideo">';
    expect(extractLinks(html)).toEqual([
      { url: 'https://t.example/video', text: 'onedriveForBusinessVideo', contentType: 'onedriveForBusinessVideo' },
    ]);
  });

  it('deduplicates the same uri matched by both attribute orders', () => {
    // The same item matches the type-first regex; the uri-first regex should
    // not add a duplicate because the uri is already seen.
    const html = '<item type="amsTranscript" uri="https://dup.example/x">';
    const links = extractLinks(html);
    const dupCount = links.filter(l => l.url === 'https://dup.example/x').length;
    expect(dupCount).toBe(1);
  });

  it('extracts both anchors and items together', () => {
    const html =
      '<a href="https://a.com">A</a><item type="t1" uri="https://i.com/1">';
    const links = extractLinks(html);
    expect(links).toContainEqual({ url: 'https://a.com', text: 'A' });
    expect(links).toContainEqual({ url: 'https://i.com/1', text: 't1', contentType: 't1' });
  });
});

describe('stripHtml', () => {
  it('removes tags and collapses whitespace', () => {
    expect(stripHtml('<p>Hello   world</p>')).toBe('Hello world');
  });

  it('decodes all supported HTML entities', () => {
    expect(stripHtml('a&nbsp;b')).toBe('a b');
    expect(stripHtml('Tom &amp; Jerry')).toBe('Tom & Jerry');
    expect(stripHtml('1 &lt; 2')).toBe('1 < 2');
    expect(stripHtml('2 &gt; 1')).toBe('2 > 1');
    expect(stripHtml('&quot;q&quot;')).toBe('"q"');
    expect(stripHtml('it&#39;s')).toBe("it's");
    expect(stripHtml('it&apos;s')).toBe("it's");
  });

  it('trims surrounding whitespace', () => {
    expect(stripHtml('   spaced   ')).toBe('spaced');
  });

  it('does not over-decode double-encoded entities', () => {
    // "&amp;lt;" is a literal ampersand-escaped "&lt;" i.e. the author wants the
    // reader to see the four characters "&lt;", not "<". Decoding &amp; before
    // &lt; would incorrectly turn this into a real "<" via a second decode pass.
    expect(stripHtml('&amp;lt;script&amp;gt;')).toBe('&lt;script&gt;');
  });

  it('returns empty string for empty input', () => {
    expect(stripHtml('')).toBe('');
  });
});

describe('escapeHtmlChars', () => {
  it('escapes ampersands first', () => {
    expect(escapeHtmlChars('a & b')).toBe('a &amp; b');
  });

  it('escapes angle brackets and quotes', () => {
    expect(escapeHtmlChars('<tag attr="v">')).toBe('&lt;tag attr=&quot;v&quot;&gt;');
  });

  it('escapes all special characters together', () => {
    expect(escapeHtmlChars('5 < 6 & "x" > 4')).toBe('5 &lt; 6 &amp; &quot;x&quot; &gt; 4');
  });

  it('leaves plain text unchanged', () => {
    expect(escapeHtmlChars('plain text 123')).toBe('plain text 123');
  });
});
