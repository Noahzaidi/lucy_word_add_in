import React, { useState, useRef, useEffect, useLayoutEffect } from 'react';
import {
  makeStyles,
  tokens,
  Badge,
  Textarea,
  Button,
  Text,
  Tooltip,
} from '@fluentui/react-components';
import {
  Send24Regular,
  Settings24Regular,
  History24Regular,
  Add24Regular,
  TextGrammarWand24Regular,
} from '@fluentui/react-icons';
import type { ChatMessage } from '../agent/orchestrator.ts';
import type { HostContext } from '../host/context.ts';
import { CrossHostBanner } from './CrossHostBanner.tsx';
import type { HostKind } from '../host/context.ts';
import { useTranslation } from '../i18n/index.ts';
import { MessageBubble } from './MessageBubble.tsx';
import { CodeBlock } from './CodeBlock.tsx';
import { CostBadge } from './CostBadge.tsx';
import type { CallCost } from '../agent/pricing.ts';

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
    justifyContent: 'space-between',
    padding: '8px 12px',
    borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
    flexShrink: 0,
  },
  brand: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    // Let the brand shrink so the action buttons are never clipped in narrow panes.
    minWidth: 0,
    overflow: 'hidden',
  },
  logo: {
    width: '24px',
    height: '24px',
    flexShrink: 0,
  },
  title: {
    fontWeight: 600,
    fontSize: '16px',
    whiteSpace: 'nowrap',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
  },
  messageList: {
    flex: 1,
    overflowY: 'auto',
    padding: '8px 0',
    display: 'flex',
    flexDirection: 'column',
    gap: '4px',
  },
  empty: {
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    flex: 1,
    gap: '8px',
    color: tokens.colorNeutralForeground3,
    padding: '24px',
    textAlign: 'center',
  },
  approvalArea: {
    padding: '8px 12px',
    borderTop: `1px solid ${tokens.colorNeutralStroke1}`,
    flexShrink: 0,
  },
  inputArea: {
    display: 'flex',
    alignItems: 'flex-end',
    gap: '8px',
    padding: '8px 12px',
    borderTop: `1px solid ${tokens.colorNeutralStroke1}`,
    flexShrink: 0,
  },
  input: {
    flex: 1,
  },
  textarea: {
    width: '100%',
    minHeight: '32px',
    maxHeight: '200px',
    resize: 'none',
    overflowY: 'auto',
  },
});

interface ChatPanelProps {
  host: HostContext;
  messages: ChatMessage[];
  isLoading: boolean;
  pendingApproval: { code: string; summary?: string } | null;
  /** Host of the currently-loaded conversation; null = no active conversation. */
  activeChatHost: HostKind | null;
  /** Running total cost for the active conversation. */
  cost: CallCost | undefined;
  /** Currently-selected provider id, used to label local-model costs as free. */
  providerId?: string;
  onSend: (text: string) => void;
  onApprove: (approved: boolean) => void;
  onOpenSettings: () => void;
  onOpenHistory: () => void;
  onOpenTighten: () => void;
  onNewChat: () => void;
}

export function ChatPanel({
  host, messages, isLoading, pendingApproval, activeChatHost, cost, providerId,
  onSend, onApprove, onOpenSettings, onOpenHistory, onOpenTighten, onNewChat,
}: ChatPanelProps) {
  const styles = useStyles();
  const { t } = useTranslation();
  const hostDisplay = t(
    host.kind === 'word' ? 'chat.hostWord' :
    host.kind === 'excel' ? 'chat.hostExcel' :
    host.kind === 'powerpoint' ? 'chat.hostPowerpoint' :
    'chat.hostOutlook',
  );
  const hostNoun = t(
    host.kind === 'word' ? 'chat.hostNounWord' :
    host.kind === 'excel' ? 'chat.hostNounExcel' :
    host.kind === 'powerpoint' ? 'chat.hostNounPowerpoint' :
    'chat.hostNounOutlook',
  );
  const headerRef = useRef<HTMLDivElement | null>(null);
  const [compactHeader, setCompactHeader] = useState(false);

  // Outlook/OWA panes can be very narrow: drop the badges instead of letting
  // them be squashed and overlap the title.
  useEffect(() => {
    const el = headerRef.current;
    if (!el || typeof ResizeObserver === 'undefined') return;
    const observer = new ResizeObserver(entries => {
      const width = entries[0]?.contentRect.width ?? 0;
      setCompactHeader(width > 0 && width < 380);
    });
    observer.observe(el);
    return () => observer.disconnect();
  }, []);

  const [inputText, setInputText] = useState('');
  const messagesEndRef = useRef<HTMLDivElement>(null);
  const textareaRef = useRef<HTMLTextAreaElement>(null);

  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [messages, pendingApproval]);

  useLayoutEffect(() => {
    const el = textareaRef.current;
    if (!el) return;
    el.style.height = 'auto';
    el.style.height = `${Math.min(el.scrollHeight, 200)}px`;
  }, [inputText]);

  const handleSubmit = () => {
    if (!inputText.trim() || isLoading) return;
    onSend(inputText);
    setInputText('');
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      handleSubmit();
    }
  };

  return (
    <div className={styles.container}>
      <div className={styles.header} ref={headerRef}>
        <div className={styles.brand}>
          <img
            src={`${import.meta.env.BASE_URL}assets/icon-64.png`}
            alt=""
            className={styles.logo}
          />
          <Text className={styles.title}>AutoOffice</Text>
          {!compactHeader && (
            <span style={{ display: 'flex', alignItems: 'center', gap: '8px', flexShrink: 0 }}>
              <Badge
                appearance="outline"
                size="small"
                color={host.kind === 'excel' ? 'success' : host.kind === 'powerpoint' ? 'danger' : host.kind === 'outlook' ? 'important' : 'brand'}
              >
                {host.displayName}
              </Badge>
              <CostBadge cost={cost} providerId={providerId} />
            </span>
          )}
        </div>
        <div style={{ display: 'flex', gap: '4px', flexShrink: 0 }}>
          {/* Word only. In browser dev mode the host also reports Word, but there is no Word API. */}
          {host.kind === 'word' && typeof Word !== 'undefined' && (
            <Tooltip content={t('tighten.openTooltip')} relationship="label">
              <Button appearance="subtle" icon={<TextGrammarWand24Regular />} onClick={onOpenTighten} disabled={isLoading} />
            </Tooltip>
          )}
          <Tooltip content={t('chat.historyTooltip')} relationship="label">
            <Button appearance="subtle" icon={<History24Regular />} onClick={onOpenHistory} disabled={isLoading} />
          </Tooltip>
          <Tooltip content={t('chat.newChatTooltip')} relationship="label">
            <Button appearance="subtle" icon={<Add24Regular />} onClick={onNewChat} disabled={isLoading} />
          </Tooltip>
          <Tooltip content={t('chat.settingsTooltip')} relationship="label">
            <Button appearance="subtle" icon={<Settings24Regular />} onClick={onOpenSettings} />
          </Tooltip>
        </div>
      </div>

      {activeChatHost && activeChatHost !== host.kind && (
        <CrossHostBanner chatHost={activeChatHost} currentHost={host.kind} />
      )}

      <div className={styles.messageList}>
        {messages.length === 0 && !pendingApproval ? (
          <div className={styles.empty}>
            <Text size={400} weight="semibold">{t('chat.welcomeTitle')}</Text>
            <Text size={200}>{t('chat.welcomeMessage', { host: hostDisplay, noun: hostNoun })}</Text>
            <Text size={200}>
              {host.kind === 'word'
                ? t('chat.exampleWord')
                : host.kind === 'excel'
                  ? t('chat.exampleExcel')
                  : host.kind === 'powerpoint'
                    ? t('chat.examplePowerpoint')
                    : t('chat.exampleOutlook')}
            </Text>
          </div>
        ) : (
          messages.map((msg, i) => (
            <MessageBubble key={i} message={msg} />
          ))
        )}
        <div ref={messagesEndRef} />
      </div>

      {pendingApproval && (
        <div className={styles.approvalArea}>
          <CodeBlock
            code={pendingApproval.code}
            summary={pendingApproval.summary}
            status="pending"
            onApprove={() => onApprove(true)}
            onReject={() => onApprove(false)}
          />
        </div>
      )}

      <div className={styles.inputArea}>
        <Textarea
          className={styles.input}
          textarea={{ ref: textareaRef, className: styles.textarea }}
          placeholder={t('chat.inputPlaceholder', { noun: hostNoun })}
          value={inputText}
          onChange={(_, data) => setInputText(data.value)}
          onKeyDown={handleKeyDown}
          disabled={isLoading}
          rows={1}
        />
        <Button
          appearance="primary"
          icon={<Send24Regular />}
          aria-label={t('chat.sendButton')}
          onClick={handleSubmit}
          disabled={!inputText.trim() || isLoading}
        />
      </div>
    </div>
  );
}
