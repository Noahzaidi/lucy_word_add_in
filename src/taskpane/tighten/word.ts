// The only Office.js code in Tighten: read the selected paragraphs, and write
// accepted text back. The model never touches the document; every write goes
// through applyRewrites, which re-checks the document first.
import { findSkipReason, splitIndent, type ParagraphFacts, type SkipReason } from './skip-rules.ts';

export const MAX_PARAGRAPHS = 30;

export interface TightenParagraph {
  /** Position in the selection; also the id the model sees. */
  id: number;
  /** The paragraph's full text, exactly as read. */
  text: string;
  /** Leading tabs/spaces, kept out of the rewrite. */
  indent: string;
  /** The text after the indent: what the model rewrites. */
  body: string;
  skipReason: SkipReason | null;
}

export type ReadResult =
  | { status: 'ok'; paragraphs: TightenParagraph[] }
  | { status: 'too-many'; count: number };

export type ApplyResult =
  | { status: 'applied'; textsAfter: string[] }
  | { status: 'changed' };

const WORD_FONT_PROPERTIES = [
  'font/bold',
  'font/italic',
  'font/underline',
  'font/strikeThrough',
  'font/doubleStrikeThrough',
  'font/superscript',
  'font/subscript',
  'font/color',
  'font/highlightColor',
  'font/size',
  'font/name',
].join(',');

/** Pure: does the document still hold exactly the paragraphs we previewed? */
export function stillMatches(expectedTexts: string[], currentTexts: string[]): boolean {
  return (
    expectedTexts.length === currentTexts.length &&
    expectedTexts.every((text, i) => text === currentTexts[i])
  );
}

/**
 * Reads the selection's paragraphs and the facts the skip rules need, in two
 * syncs. Stops after the first sync when more than MAX_PARAGRAPHS are selected.
 */
async function readSelectedParagraphs(
  context: Word.RequestContext,
): Promise<{ paragraphs: Word.Paragraph[]; facts: ParagraphFacts[] } | { tooMany: number }> {
  const collection = context.document.getSelection().paragraphs;
  collection.load('text');
  await context.sync();
  if (collection.items.length > MAX_PARAGRAPHS) {
    return { tooMany: collection.items.length };
  }

  const loads = collection.items.map((paragraph) => ({
    pictures: paragraph.inlinePictures.load('$none'),
    links: paragraph.getRange().getHyperlinkRanges().load('$none'),
    contentControls: paragraph.contentControls.load('$none'),
    words: paragraph.getTextRanges([' '], true).load(WORD_FONT_PROPERTIES),
  }));
  await context.sync();

  const facts = collection.items.map((paragraph, i) => ({
    text: paragraph.text,
    pictureCount: loads[i].pictures.items.length,
    linkCount: loads[i].links.items.length,
    contentControlCount: loads[i].contentControls.items.length,
    wordFonts: loads[i].words.items.map(({ font }) => ({
      bold: font.bold,
      italic: font.italic,
      underline: font.underline,
      strikeThrough: font.strikeThrough,
      doubleStrikeThrough: font.doubleStrikeThrough,
      superscript: font.superscript,
      subscript: font.subscript,
      color: font.color,
      highlightColor: font.highlightColor,
      size: font.size,
      name: font.name,
    })),
  }));
  return { paragraphs: collection.items, facts };
}

export async function readSelection(): Promise<ReadResult> {
  return Word.run(async (context) => {
    const read = await readSelectedParagraphs(context);
    if ('tooMany' in read) {
      return { status: 'too-many', count: read.tooMany };
    }
    const paragraphs = read.facts.map((facts, id) => ({
      id,
      text: facts.text,
      ...splitIndent(facts.text),
      skipReason: findSkipReason(facts),
    }));
    return { status: 'ok', paragraphs };
  });
}

/**
 * Writes newTexts (paragraph id -> full new text) only if the selection still
 * holds exactly expectedTexts and every target is still safe to rewrite.
 * Otherwise it writes nothing. Used for Accept and, with the roles swapped, Undo.
 */
export async function applyRewrites(
  expectedTexts: string[],
  newTexts: Map<number, string>,
): Promise<ApplyResult> {
  return Word.run(async (context) => {
    const read = await readSelectedParagraphs(context);
    if ('tooMany' in read) return { status: 'changed' };

    const currentTexts = read.facts.map((facts) => facts.text);
    if (!stillMatches(expectedTexts, currentTexts)) return { status: 'changed' };
    for (const id of newTexts.keys()) {
      if (findSkipReason(read.facts[id]) !== null) return { status: 'changed' };
    }

    // 'Content' is the paragraph without its end-of-paragraph mark, which holds
    // the style and list membership, so headings and list items keep them.
    for (const [id, text] of newTexts) {
      read.paragraphs[id].getRange('Content').insertText(text, 'Replace');
    }
    await context.sync();

    // What the document should now hold at the same positions. Undo expects
    // exactly this, so if the selection moved or Word altered the text, Undo
    // refuses rather than writing to whatever paragraph is now in that position.
    const textsAfter = expectedTexts.map((text, id) => newTexts.get(id) ?? text);
    return { status: 'applied', textsAfter };
  });
}
