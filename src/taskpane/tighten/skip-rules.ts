// Decides which selected paragraphs Tighten may rewrite. Tighten writes plain
// text, so a paragraph is only eligible when none of the known cases a plain-text
// replace would lose is present: images, links, content controls, mixed
// formatting, and line breaks or other special characters. Fields, comments,
// notes, bookmarks and tracked changes are not detected (see NOTES.md).
// word.ts collects these facts from Word; this file only decides.

/** Font properties of one word, as Word reports them. */
export interface WordFont {
  bold: boolean | null;
  italic: boolean | null;
  underline: string | null;
  strikeThrough: boolean | null;
  doubleStrikeThrough: boolean | null;
  superscript: boolean | null;
  subscript: boolean | null;
  color: string | null;
  highlightColor: string | null;
  size: number | null;
  name: string | null;
}

export interface ParagraphFacts {
  text: string;
  pictureCount: number;
  linkCount: number;
  contentControlCount: number;
  wordFonts: WordFont[];
}

export type SkipReason =
  | 'empty'
  | 'special-characters'
  | 'image'
  | 'link'
  | 'content-control'
  | 'mixed-formatting';

/**
 * Line breaks (Shift+Enter is \v in Word), tabs, page breaks and other control
 * characters. A plain-text write drops or mangles them, so a paragraph that
 * contains one is skipped, and a rewrite that contains one is rejected.
 */
export const CONTROL_CHARACTERS = /[\u0000-\u001f\u007f\u2028\u2029]/;

/** Leading tabs/spaces are indentation: kept as-is and not sent to the model. */
export function splitIndent(text: string): { indent: string; body: string } {
  const indent = text.match(/^[\t ]*/)?.[0] ?? '';
  return { indent, body: text.slice(indent.length) };
}

/**
 * True when formatting varies inside a single word (e.g. a superscript "st" in
 * "1st"). Word documents two markers: highlightColor is '' when mixed (null just
 * means "no highlight") and underline is 'Mixed'. For the other properties it
 * documents nothing, so a null or empty value is treated as mixed: skipping is
 * the safe side.
 */
function variesInsideWord(font: WordFont): boolean {
  const isUnknown = (value: unknown) => value === null || value === '';
  return (
    isUnknown(font.bold) ||
    isUnknown(font.italic) ||
    isUnknown(font.strikeThrough) ||
    isUnknown(font.doubleStrikeThrough) ||
    isUnknown(font.superscript) ||
    isUnknown(font.subscript) ||
    isUnknown(font.size) ||
    isUnknown(font.color) ||
    isUnknown(font.name) ||
    isUnknown(font.underline) ||
    font.underline === 'Mixed' ||
    font.highlightColor === ''
  );
}

function hasMixedFormatting(fonts: WordFont[]): boolean {
  // No word ranges for a non-empty paragraph means we couldn't check: skip it.
  if (fonts.length === 0) return true;
  const distinctFormats = new Set(fonts.map((font) => JSON.stringify(font)));
  return distinctFormats.size > 1 || fonts.some(variesInsideWord);
}

export function findSkipReason(facts: ParagraphFacts): SkipReason | null {
  const { body } = splitIndent(facts.text);
  if (body.trim() === '') return 'empty';
  if (CONTROL_CHARACTERS.test(body)) return 'special-characters';
  if (facts.pictureCount > 0) return 'image';
  if (facts.linkCount > 0) return 'link';
  if (facts.contentControlCount > 0) return 'content-control';
  if (hasMixedFormatting(facts.wordFonts)) return 'mixed-formatting';
  return null;
}
