import { describe, it, expect } from 'vitest';
import { validateRewrites, TIGHTEN_SCHEMA } from './validate.ts';

const sent = [
  { id: 0, text: 'In order to make sure that we are able to finish on time, we should start now.' },
  { id: 2, text: 'Quarterly results' },
];

describe('validateRewrites', () => {
  it('accepts exactly one single-line rewrite per paragraph sent', () => {
    const result = validateRewrites(sent, {
      paragraphs: [
        { id: 0, text: '  To finish on time, we should start now. ' },
        { id: 2, text: 'Quarterly results' },
      ],
    });

    expect(result).toEqual({
      ok: true,
      rewrites: new Map([
        [0, 'To finish on time, we should start now.'],
        [2, 'Quarterly results'],
      ]),
    });
  });

  it('rejects a missing paragraph', () => {
    const result = validateRewrites(sent, { paragraphs: [{ id: 0, text: 'Start now.' }] });
    expect(result.ok).toBe(false);
  });

  it('rejects an id that was never sent', () => {
    const result = validateRewrites(sent, {
      paragraphs: [{ id: 0, text: 'Start now.' }, { id: 7, text: 'Injected paragraph' }],
    });
    expect(result).toEqual({ ok: false, reason: 'Unknown paragraph id 7.' });
  });

  it('rejects a duplicated id', () => {
    const result = validateRewrites(sent, {
      paragraphs: [{ id: 0, text: 'Start now.' }, { id: 0, text: 'Start now!' }],
    });
    expect(result).toEqual({ ok: false, reason: 'Paragraph 0 appears twice.' });
  });

  it('rejects an empty rewrite', () => {
    const result = validateRewrites(sent, {
      paragraphs: [{ id: 0, text: '   ' }, { id: 2, text: 'Quarterly results' }],
    });
    expect(result.ok).toBe(false);
  });

  it('rejects line breaks, which would split the paragraph', () => {
    const result = validateRewrites(sent, {
      paragraphs: [{ id: 0, text: 'Start now.\nOr tomorrow.' }, { id: 2, text: 'Quarterly results' }],
    });
    expect(result.ok).toBe(false);
  });

  it('rejects a rewrite far longer than the original', () => {
    const result = validateRewrites(sent, {
      paragraphs: [{ id: 0, text: 'Start now.' }, { id: 2, text: 'x'.repeat(200) }],
    });
    expect(result.ok).toBe(false);
  });
});

describe('TIGHTEN_SCHEMA', () => {
  it('rejects replies that are not the expected shape', () => {
    expect(TIGHTEN_SCHEMA.safeParse({ paragraphs: [{ id: '0', text: 'x' }] }).success).toBe(false);
    expect(TIGHTEN_SCHEMA.safeParse({ rewrites: [] }).success).toBe(false);
  });
});
