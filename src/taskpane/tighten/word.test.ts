import { describe, it, expect, vi, afterEach } from 'vitest';
import { stillMatches, applyRewrites } from './word.ts';
import type { WordFont } from './skip-rules.ts';

// stillMatches is the guard Accept and Undo rely on: write only if the selection
// still holds exactly the paragraphs that were previewed.
describe('stillMatches', () => {
  const previewed = ['Title', '\tFirst paragraph.', '', 'Bullet one'];

  it('matches the same paragraphs with the same text', () => {
    expect(stillMatches(previewed, [...previewed])).toBe(true);
  });

  it('refuses when a paragraph was edited after the preview', () => {
    expect(stillMatches(previewed, ['Title', '\tFirst paragraph, edited.', '', 'Bullet one'])).toBe(false);
  });

  it('refuses when the selection now covers more or fewer paragraphs', () => {
    expect(stillMatches(previewed, previewed.slice(0, 3))).toBe(false);
    expect(stillMatches(previewed, [...previewed, 'Bullet two'])).toBe(false);
  });

  it('refuses even a whitespace-only difference', () => {
    expect(stillMatches(previewed, ['Title ', '\tFirst paragraph.', '', 'Bullet one'])).toBe(false);
  });
});

// A minimal stand-in for the parts of Word the Tighten code touches, so the
// write guard can be tested without Office.
const plainFont: WordFont = {
  bold: false,
  italic: false,
  underline: 'None',
  strikeThrough: false,
  doubleStrikeThrough: false,
  superscript: false,
  subscript: false,
  color: '#000000',
  highlightColor: null,
  size: 11,
  name: 'Aptos',
};

function fakeParagraph(text: string, linkCount = 0) {
  const contentRange = { insertText: vi.fn() };
  const collection = (count: number) => ({ load: () => ({ items: new Array(count).fill({}) }) });
  return {
    text,
    contentRange,
    inlinePictures: collection(0),
    contentControls: collection(0),
    getTextRanges: () => ({ load: () => ({ items: [{ font: plainFont }] }) }),
    getRange: (location?: string) =>
      location === 'Content' ? contentRange : { getHyperlinkRanges: () => collection(linkCount) },
  };
}

function stubWordSelection(paragraphs: ReturnType<typeof fakeParagraph>[]) {
  const context = {
    document: { getSelection: () => ({ paragraphs: { items: paragraphs, load: () => {} } }) },
    sync: async () => {},
  };
  vi.stubGlobal('Word', { run: (batch: (ctx: unknown) => unknown) => batch(context) });
}

describe('applyRewrites', () => {
  afterEach(() => vi.unstubAllGlobals());

  it('writes only the changed paragraphs, through their Content range', async () => {
    const paragraphs = [fakeParagraph('Title'), fakeParagraph('In order to start, we need to begin.')];
    stubWordSelection(paragraphs);

    const result = await applyRewrites(['Title', 'In order to start, we need to begin.'], new Map([[1, 'To start, begin.']]));

    expect(paragraphs[0].contentRange.insertText).not.toHaveBeenCalled();
    expect(paragraphs[1].contentRange.insertText).toHaveBeenCalledWith('To start, begin.', 'Replace');
    // Undo will expect exactly this.
    expect(result).toEqual({ status: 'applied', textsAfter: ['Title', 'To start, begin.'] });
  });

  it('writes nothing when a paragraph was edited after the preview', async () => {
    const paragraphs = [fakeParagraph('Title'), fakeParagraph('Edited by the user.')];
    stubWordSelection(paragraphs);

    const result = await applyRewrites(['Title', 'Original text.'], new Map([[1, 'Shorter.']]));

    expect(result).toEqual({ status: 'changed' });
    expect(paragraphs[1].contentRange.insertText).not.toHaveBeenCalled();
  });

  it('writes nothing when the selection now covers different paragraphs', async () => {
    const paragraphs = [fakeParagraph('Title')];
    stubWordSelection(paragraphs);

    const result = await applyRewrites(['Title', 'Original text.'], new Map([[1, 'Shorter.']]));

    expect(result).toEqual({ status: 'changed' });
    expect(paragraphs[0].contentRange.insertText).not.toHaveBeenCalled();
  });

  it('writes nothing when a target paragraph gained a link after the preview', async () => {
    const paragraphs = [fakeParagraph('See the policy page.', 1)];
    stubWordSelection(paragraphs);

    const result = await applyRewrites(['See the policy page.'], new Map([[0, 'See the policy.']]));

    expect(result).toEqual({ status: 'changed' });
    expect(paragraphs[0].contentRange.insertText).not.toHaveBeenCalled();
  });

  it('undo is the same call in reverse: it restores the originals', async () => {
    const paragraphs = [fakeParagraph('To start, begin.')];
    stubWordSelection(paragraphs);

    const result = await applyRewrites(['To start, begin.'], new Map([[0, 'In order to start, we need to begin.']]));

    expect(result.status).toBe('applied');
    expect(paragraphs[0].contentRange.insertText).toHaveBeenCalledWith('In order to start, we need to begin.', 'Replace');
  });
});
