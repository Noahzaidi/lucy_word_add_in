import React, { useState } from 'react';
import { makeStyles, tokens, Button, Text, Spinner } from '@fluentui/react-components';
import {
  ArrowLeft24Regular,
  ArrowUndo24Regular,
  Checkmark24Regular,
  Dismiss24Regular,
  TextGrammarWand24Regular,
} from '@fluentui/react-icons';
import type { AppSettings } from '../store/settings.ts';
import { formatError, type FormattedError } from '../agent/errors.ts';
import { useTranslation, useFormatters, type TranslationKey } from '../i18n/index.ts';
import { ErrorBubble } from './ErrorBubble.tsx';
import { readSelection, applyRewrites, MAX_PARAGRAPHS, type TightenParagraph } from '../tighten/word.ts';
import { requestRewrites } from '../tighten/rewrite.ts';
import { validateRewrites } from '../tighten/validate.ts';
import type { SkipReason } from '../tighten/skip-rules.ts';

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    height: '100%',
    overflow: 'hidden',
  },
  header: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    padding: '8px 12px',
    borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
    flexShrink: 0,
  },
  content: {
    flex: 1,
    overflowY: 'auto',
    display: 'flex',
    flexDirection: 'column',
    gap: '12px',
    padding: '12px',
  },
  hint: {
    color: tokens.colorNeutralForeground2,
  },
  card: {
    backgroundColor: tokens.colorNeutralBackground3,
    border: `1px solid ${tokens.colorNeutralStroke1}`,
    borderRadius: '8px',
    padding: '8px 12px',
    display: 'flex',
    flexDirection: 'column',
    gap: '4px',
  },
  label: {
    fontSize: '11px',
    fontWeight: 600,
    textTransform: 'uppercase',
    color: tokens.colorNeutralForeground3,
  },
  before: {
    color: tokens.colorNeutralForeground3,
    whiteSpace: 'pre-wrap',
  },
  after: {
    whiteSpace: 'pre-wrap',
  },
  actions: {
    display: 'flex',
    gap: '8px',
  },
});

const SKIP_REASON_KEYS: Record<Exclude<SkipReason, 'empty'>, TranslationKey> = {
  'special-characters': 'tighten.skipSpecialCharacters',
  image: 'tighten.skipImage',
  link: 'tighten.skipLink',
  'content-control': 'tighten.skipContentControl',
  'mixed-formatting': 'tighten.skipMixedFormatting',
};

/** What Undo needs: the text Accept left in the document, and the originals to restore. */
interface UndoData {
  textsAfter: string[];
  originals: Map<number, string>;
}

type PanelState =
  | { step: 'idle' }
  | { step: 'working' }
  // changes: paragraph id -> rewritten text, only for paragraphs the model actually changed.
  | { step: 'preview'; paragraphs: TightenParagraph[]; changes: Map<number, string> }
  // skipped: shown when nothing could be tightened, so the user sees why.
  | { step: 'result'; message: string; undo?: UndoData; skipped?: TightenParagraph[] }
  | { step: 'error'; error: FormattedError };

interface TightenPanelProps {
  settings: AppSettings;
  onClose: () => void;
}

export function TightenPanel({ settings, onClose }: TightenPanelProps) {
  const styles = useStyles();
  const { t } = useTranslation();
  const { formatPlural } = useFormatters();
  const [state, setState] = useState<PanelState>({ step: 'idle' });

  const showResult = (message: string, undo?: UndoData) => setState({ step: 'result', message, undo });

  // Read the selection, ask the model, validate the reply, then show the preview.
  // Nothing is written to the document here.
  const handleTighten = async () => {
    setState({ step: 'working' });
    try {
      const selection = await readSelection();
      if (selection.status === 'too-many') {
        showResult(t('tighten.tooMany', { max: MAX_PARAGRAPHS, count: selection.count }));
        return;
      }

      const sent = selection.paragraphs
        .filter((paragraph) => paragraph.skipReason === null)
        .map((paragraph) => ({ id: paragraph.id, text: paragraph.body }));
      if (sent.length === 0) {
        setState({ step: 'result', message: t('tighten.nothingToTighten'), skipped: selection.paragraphs });
        return;
      }

      const reply = await requestRewrites(sent, settings);
      if (!reply.ok) {
        setState({ step: 'error', error: reply.error });
        return;
      }

      const validation = validateRewrites(sent, reply.output);
      if (!validation.ok) {
        setState({
          step: 'error',
          error: { kind: 'api', title: t('tighten.invalidReplyTitle'), detail: `${validation.reason} ${t('tighten.nothingChanged')}` },
        });
        return;
      }

      const changes = new Map<number, string>();
      for (const paragraph of selection.paragraphs) {
        const rewrite = validation.rewrites.get(paragraph.id);
        if (rewrite !== undefined && rewrite !== paragraph.body.trim()) changes.set(paragraph.id, rewrite);
      }
      if (changes.size === 0) {
        showResult(t('tighten.noSuggestions'));
        return;
      }

      setState({ step: 'preview', paragraphs: selection.paragraphs, changes });
    } catch (err) {
      setState({ step: 'error', error: formatError(err) });
    }
  };

  // The only path that writes: applyRewrites re-checks the document first.
  const handleAccept = async (paragraphs: TightenParagraph[], changes: Map<number, string>) => {
    const newTexts = new Map<number, string>();
    const originals = new Map<number, string>();
    for (const paragraph of paragraphs) {
      const rewrite = changes.get(paragraph.id);
      if (rewrite !== undefined) {
        newTexts.set(paragraph.id, paragraph.indent + rewrite);
        originals.set(paragraph.id, paragraph.text);
      }
    }

    setState({ step: 'working' });
    try {
      const result = await applyRewrites(paragraphs.map((paragraph) => paragraph.text), newTexts);
      if (result.status === 'changed') {
        showResult(t('tighten.documentChanged'));
        return;
      }
      const message = formatPlural(newTexts.size, { one: t('tighten.applied_one'), other: t('tighten.applied_other') });
      showResult(message, { textsAfter: result.textsAfter, originals });
    } catch (err) {
      // Office.js batches are not transactions: some paragraphs may already be written.
      const error = formatError(err);
      setState({ step: 'error', error: { ...error, detail: `${error.detail} ${t('tighten.partialWriteHint')}` } });
    }
  };

  // Undo is Accept in reverse: same checks, writing the original text back.
  const handleUndo = async (undo: UndoData) => {
    setState({ step: 'working' });
    try {
      const result = await applyRewrites(undo.textsAfter, undo.originals);
      showResult(t(result.status === 'applied' ? 'tighten.undone' : 'tighten.undoRefused'));
    } catch (err) {
      setState({ step: 'error', error: formatError(err) });
    }
  };

  const renderCard = (paragraph: TightenParagraph, changes: Map<number, string>) => {
    const rewrite = changes.get(paragraph.id);
    if (paragraph.skipReason !== null && paragraph.skipReason !== 'empty') {
      return (
        <div key={paragraph.id} className={styles.card}>
          <Text className={styles.label}>{t('tighten.skipped', { reason: t(SKIP_REASON_KEYS[paragraph.skipReason]) })}</Text>
          <Text className={styles.before}>{paragraph.body}</Text>
        </div>
      );
    }
    if (rewrite === undefined) {
      return (
        <div key={paragraph.id} className={styles.card}>
          <Text className={styles.label}>{t('tighten.noChange')}</Text>
          <Text className={styles.after}>{paragraph.body}</Text>
        </div>
      );
    }
    return (
      <div key={paragraph.id} className={styles.card}>
        <Text className={styles.label}>{t('tighten.before')}</Text>
        <Text className={styles.before}>{paragraph.body}</Text>
        <Text className={styles.label}>{t('tighten.after')}</Text>
        <Text className={styles.after}>{rewrite}</Text>
      </div>
    );
  };

  return (
    <div className={styles.container}>
      <div className={styles.header}>
        <Button appearance="subtle" icon={<ArrowLeft24Regular />} onClick={onClose} aria-label={t('settings.backButton')} />
        <Text weight="semibold" size={400}>{t('tighten.title')}</Text>
      </div>

      <div className={styles.content}>
        {state.step === 'preview' ? (
          <>
            {state.paragraphs
              .filter((paragraph) => paragraph.skipReason !== 'empty')
              .map((paragraph) => renderCard(paragraph, state.changes))}
            <div className={styles.actions}>
              <Button
                appearance="primary"
                icon={<Checkmark24Regular />}
                onClick={() => handleAccept(state.paragraphs, state.changes)}
              >
                {t('tighten.acceptButton')}
              </Button>
              <Button icon={<Dismiss24Regular />} onClick={() => showResult(t('tighten.rejected'))}>
                {t('tighten.rejectButton')}
              </Button>
            </div>
          </>
        ) : (
          <>
            <Text className={styles.hint}>{t('tighten.hint')}</Text>
            <div className={styles.actions}>
              <Button
                appearance="primary"
                icon={<TextGrammarWand24Regular />}
                onClick={handleTighten}
                disabled={state.step === 'working'}
              >
                {t('tighten.runButton')}
              </Button>
            </div>
          </>
        )}

        {state.step === 'working' && <Spinner size="small" label={t('tighten.working')} />}

        {state.step === 'result' && (
          <>
            <Text>{state.message}</Text>
            {state.skipped
              ?.filter((paragraph) => paragraph.skipReason !== 'empty')
              .map((paragraph) => renderCard(paragraph, new Map()))}
            {state.undo && (
              <div className={styles.actions}>
                <Button icon={<ArrowUndo24Regular />} onClick={() => handleUndo(state.undo!)}>
                  {t('tighten.undoButton')}
                </Button>
              </div>
            )}
          </>
        )}

        {state.step === 'error' && <ErrorBubble {...state.error} />}
      </div>
    </div>
  );
}
