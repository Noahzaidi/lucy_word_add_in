import { z } from 'zod';
import { CONTROL_CHARACTERS } from './skip-rules.ts';

/** The structured output we ask the model for: one entry per paragraph, matched by id. */
export const TIGHTEN_SCHEMA = z.object({
  paragraphs: z.array(z.object({ id: z.number(), text: z.string() })),
});

export type TightenOutput = z.infer<typeof TIGHTEN_SCHEMA>;

/** A paragraph as sent to the model. */
export interface SentParagraph {
  id: number;
  text: string;
}

export type ValidationResult =
  | { ok: true; rewrites: Map<number, string> }
  | { ok: false; reason: string };

// A tightened paragraph should be shorter. This bound only stops a runaway or
// hijacked reply from pouring a wall of text into one paragraph.
const MAX_LENGTH_FACTOR = 2;
const MAX_LENGTH_SLACK = 50;

/**
 * Checks the model's reply against what was sent, before anything reaches the
 * preview: exactly one single-line, non-empty rewrite per paragraph sent.
 */
export function validateRewrites(sent: SentParagraph[], output: TightenOutput): ValidationResult {
  if (output.paragraphs.length !== sent.length) {
    return { ok: false, reason: `Expected ${sent.length} paragraphs, got ${output.paragraphs.length}.` };
  }

  const originals = new Map(sent.map((paragraph) => [paragraph.id, paragraph.text]));
  const rewrites = new Map<number, string>();

  for (const entry of output.paragraphs) {
    const original = originals.get(entry.id);
    if (original === undefined) {
      return { ok: false, reason: `Unknown paragraph id ${entry.id}.` };
    }
    if (rewrites.has(entry.id)) {
      return { ok: false, reason: `Paragraph ${entry.id} appears twice.` };
    }

    const text = entry.text.trim();
    if (text === '') {
      return { ok: false, reason: `Paragraph ${entry.id} came back empty.` };
    }
    if (CONTROL_CHARACTERS.test(text)) {
      return { ok: false, reason: `Paragraph ${entry.id} contains a line break or control character.` };
    }
    if (text.length > original.length * MAX_LENGTH_FACTOR + MAX_LENGTH_SLACK) {
      return { ok: false, reason: `Paragraph ${entry.id} came back much longer than the original.` };
    }
    rewrites.set(entry.id, text);
  }

  return { ok: true, rewrites };
}
