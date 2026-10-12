import { describe, it, expect } from 'vitest';
import { findSkipReason, splitIndent, type ParagraphFacts, type WordFont } from './skip-rules.ts';

// highlightColor null is how Word reports "no highlight"; it must not count as mixed.
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

function facts(overrides: Partial<ParagraphFacts> = {}): ParagraphFacts {
  return {
    text: 'A plain paragraph of text.',
    pictureCount: 0,
    linkCount: 0,
    contentControlCount: 0,
    wordFonts: [plainFont, plainFont, plainFont],
    ...overrides,
  };
}

describe('findSkipReason', () => {
  it('accepts a plain paragraph with uniform formatting', () => {
    expect(findSkipReason(facts())).toBeNull();
  });

  it('accepts a uniformly bold paragraph (only mixed formatting is a problem)', () => {
    const bold = { ...plainFont, bold: true };
    expect(findSkipReason(facts({ wordFonts: [bold, bold] }))).toBeNull();
  });

  it('accepts a paragraph that only starts with a tab', () => {
    expect(findSkipReason(facts({ text: '\tIndented first line.' }))).toBeNull();
  });

  it('skips empty and whitespace-only paragraphs', () => {
    expect(findSkipReason(facts({ text: '' }))).toBe('empty');
    expect(findSkipReason(facts({ text: '\t  ' }))).toBe('empty');
  });

  it('skips soft line breaks (Shift+Enter) and inner tabs', () => {
    expect(findSkipReason(facts({ text: 'TITLE 1\v\vTitle 2' }))).toBe('special-characters');
    expect(findSkipReason(facts({ text: 'Name\tValue' }))).toBe('special-characters');
  });

  it('skips images, links and content controls', () => {
    expect(findSkipReason(facts({ pictureCount: 1 }))).toBe('image');
    expect(findSkipReason(facts({ linkCount: 1 }))).toBe('link');
    expect(findSkipReason(facts({ contentControlCount: 1 }))).toBe('content-control');
  });

  it('skips a paragraph where one word is formatted differently', () => {
    const boldWord = { ...plainFont, bold: true };
    expect(findSkipReason(facts({ wordFonts: [plainFont, boldWord, plainFont] }))).toBe('mixed-formatting');
  });

  it('skips a superscript word, such as the "st" in "1st"', () => {
    const superscriptWord = { ...plainFont, superscript: true };
    expect(findSkipReason(facts({ wordFonts: [plainFont, superscriptWord] }))).toBe('mixed-formatting');
    expect(findSkipReason(facts({ wordFonts: [{ ...plainFont, superscript: null }] }))).toBe('mixed-formatting');
  });

  it('skips a non-empty paragraph whose words could not be read', () => {
    expect(findSkipReason(facts({ wordFonts: [] }))).toBe('mixed-formatting');
  });

  it('skips a paragraph where formatting varies inside one word', () => {
    expect(findSkipReason(facts({ wordFonts: [{ ...plainFont, bold: null }] }))).toBe('mixed-formatting');
    expect(findSkipReason(facts({ wordFonts: [{ ...plainFont, underline: 'Mixed' }] }))).toBe('mixed-formatting');
    expect(findSkipReason(facts({ wordFonts: [{ ...plainFont, highlightColor: '' }] }))).toBe('mixed-formatting');
  });
});

describe('splitIndent', () => {
  it('separates leading tabs and spaces from the text', () => {
    expect(splitIndent('\t  Body text')).toEqual({ indent: '\t  ', body: 'Body text' });
    expect(splitIndent('No indent')).toEqual({ indent: '', body: 'No indent' });
  });
});
