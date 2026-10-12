import React, { useState, useEffect, useRef, useCallback } from 'react';
import {
  makeStyles,
  tokens,
  Button,
  Input,
  Label,
  Select,
  Switch,
  Text,
  Divider,
  Field,
  Badge,
  Spinner,
  Tooltip,
} from '@fluentui/react-components';
import {
  ArrowLeft24Regular,
  Add24Regular,
  Delete24Regular,
  Eye24Regular,
  EyeOff24Regular,
  ArrowClockwise24Regular,
} from '@fluentui/react-icons';
import type { AppSettings, McpServerConfig } from '../store/settings.ts';
import { LOCAL_PROVIDER_IDS } from '../store/settings.ts';
import { useTranslation, availableLocales, type LocaleId, type TranslationKey } from '../i18n/index.ts';
import { discoverLmStudioModels } from '../agent/lmstudio.ts';
import { DEFAULT_OPENWEBUI_BASE_URL, discoverOpenWebUiModels } from '../agent/openwebui.ts';

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
    padding: '16px 12px',
    display: 'flex',
    flexDirection: 'column',
    gap: '16px',
  },
  section: {
    display: 'flex',
    flexDirection: 'column',
    gap: '8px',
  },
  row: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
  },
  mcpServer: {
    display: 'flex',
    flexDirection: 'column',
    gap: '6px',
    padding: '8px',
    backgroundColor: tokens.colorNeutralBackground3,
    borderRadius: '6px',
  },
  keyInput: {
    flex: 1,
  },
  statusRow: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    minHeight: '24px',
  },
});

interface SettingsPanelProps {
  settings: AppSettings;
  onChange: (settings: AppSettings) => void;
  onClose: () => void;
}

const PROVIDER_MODELS: Record<string, string[]> = {
  anthropic: [
    'claude-opus-4-7',
    'claude-sonnet-4-6',
    'claude-haiku-4-5-20251001',
    'claude-opus-4-6',
  ],
  openai: [
    'gpt-5.4',
    'gpt-5.4-pro',
    'gpt-5.4-mini',
    'gpt-5.4-nano',
    'gpt-5.3-chat-latest',
    'gpt-5.3-codex',
    'gpt-5',
    'gpt-4o',
    'gpt-4o-mini',
  ],
  google: [
    'gemini-3.1-pro-preview',
    'gemini-3-flash-preview',
    'gemini-3.1-flash-lite-preview',
    'gemini-2.5-pro',
    'gemini-2.5-flash',
    'gemini-2.5-flash-lite',
  ],
  groq: [
    'moonshotai/kimi-k2-instruct-0905',
    'meta-llama/llama-4-scout-17b-16e-instruct',
    'llama-3.3-70b-versatile',
    'llama-3.1-8b-instant',
    'qwen/qwen3-32b',
    'openai/gpt-oss-120b',
    'openai/gpt-oss-20b',
  ],
  xai: [
    'grok-4',
    'grok-4-fast',
    'grok-code-fast-1',
    'grok-3',
    'grok-3-mini',
  ],
  deepseek: [
    'deepseek-chat',
    'deepseek-reasoner',
  ],
  gateway: [],
  openrouter: [],
  ollama: [],
  lmstudio: [],
  openwebui: [],
  'openai-compatible': [],
};

const PROVIDERS_WITH_BASE_URL = new Set(['lucy-gateway', 'openai-compatible', 'openrouter', 'ollama', 'lmstudio', 'openwebui']);

/** Providers whose model list is discovered from the server they point at. */
const PROVIDERS_WITH_MODEL_DISCOVERY = new Set(['lmstudio', 'openwebui']);

const DEFAULT_DISCOVERY_BASE_URLS: Record<string, string> = {
  lmstudio: 'http://localhost:1234/v1',
  openwebui: DEFAULT_OPENWEBUI_BASE_URL,
};

/** Self-hosted providers where an API key is optional but still accepted. */
const PROVIDERS_WITH_OPTIONAL_API_KEY = new Set(['openwebui']);

interface DiscoveredModel {
  id: string;
  displayName?: string;
  state?: 'loaded' | 'not-loaded';
}

interface DiscoveredModelList {
  status: 'idle' | 'connecting' | 'connected' | 'unreachable';
  models: DiscoveredModel[];
  error?: string;
}

interface ModelDiscoveryOptions {
  signal?: AbortSignal;
  force?: boolean;
  apiKey?: string;
}

function discoverProviderModels(
  providerId: string,
  baseUrl: string,
  opts: ModelDiscoveryOptions
): Promise<DiscoveredModelList> {
  return providerId === 'openwebui'
    ? discoverOpenWebUiModels(baseUrl, opts)
    : discoverLmStudioModels(baseUrl, opts);
}

export function SettingsPanel({ settings, onChange, onClose }: SettingsPanelProps) {
  const styles = useStyles();
  const { t, locale, setLocale } = useTranslation();
  const [showKeys, setShowKeys] = useState<Record<string, boolean>>({});

  const updateProvider = (id: string, field: string, value: string) => {
    const providers = settings.providers.map(p =>
      p.id === id ? { ...p, [field]: value } : p
    );
    onChange({ ...settings, providers });
  };

  const addMcpServer = () => {
    const newServer: McpServerConfig = {
      name: `Server ${settings.mcpServers.length + 1}`,
      url: '',
      transport: 'streamable-http',
      enabled: true,
    };
    onChange({ ...settings, mcpServers: [...settings.mcpServers, newServer] });
  };

  const updateMcpServer = (index: number, updates: Partial<McpServerConfig>) => {
    const servers = settings.mcpServers.map((s, i) =>
      i === index ? { ...s, ...updates } : s
    );
    onChange({ ...settings, mcpServers: servers });
  };

  const removeMcpServer = (index: number) => {
    onChange({
      ...settings,
      mcpServers: settings.mcpServers.filter((_, i) => i !== index),
    });
  };

  const selectedProvider = settings.providers.find(p => p.id === settings.selectedProviderId);
  const models = PROVIDER_MODELS[settings.selectedProviderId] || [];
  // Lucy Gateway's key field holds the gateway token, not a model API key.
  const isLucyGateway = settings.selectedProviderId === 'lucy-gateway';

  return (
    <div className={styles.container}>
      <div className={styles.header}>
        <Button appearance="subtle" icon={<ArrowLeft24Regular />} onClick={onClose} />
        <Text weight="semibold" size={400}>{t('settings.title')}</Text>
      </div>

      <div className={styles.content}>
        {/* AI Provider */}
        <div className={styles.section}>
          <Text weight="semibold" size={300}>{t('settings.providerSection')}</Text>

          <Field label={t('settings.providerLabel')}>
            <Select
              value={settings.selectedProviderId}
              onChange={(_, data) => onChange({ ...settings, selectedProviderId: data.value, selectedModel: '' })}
            >
              <option value="">{t('settings.providerPlaceholder')}</option>
              {settings.providers.map(p => (
                <option key={p.id} value={p.id}>{p.name}</option>
              ))}
            </Select>
          </Field>

          {selectedProvider && (
            <>
              {(!LOCAL_PROVIDER_IDS.has(selectedProvider.id) ||
                PROVIDERS_WITH_OPTIONAL_API_KEY.has(selectedProvider.id)) && (
                <Field label={isLucyGateway ? t('settings.gatewayTokenLabel') : t('settings.apiKeyLabel')}>
                  <div className={styles.row}>
                    <Input
                      className={styles.keyInput}
                      type={showKeys[selectedProvider.id] ? 'text' : 'password'}
                      value={selectedProvider.apiKey}
                      onChange={(_, data) => updateProvider(selectedProvider.id, 'apiKey', data.value)}
                      placeholder={isLucyGateway ? t('settings.gatewayTokenPlaceholder') : t('settings.apiKeyPlaceholder')}
                    />
                    <Button
                      appearance="subtle"
                      icon={showKeys[selectedProvider.id] ? <EyeOff24Regular /> : <Eye24Regular />}
                      onClick={() => setShowKeys(prev => ({ ...prev, [selectedProvider.id]: !prev[selectedProvider.id] }))}
                    />
                  </div>
                </Field>
              )}

              {PROVIDERS_WITH_BASE_URL.has(selectedProvider.id) && (
                <Field label={t('settings.baseUrlLabel')}>
                  <Input
                    value={selectedProvider.baseUrl || ''}
                    onChange={(_, data) => updateProvider(selectedProvider.id, 'baseUrl', data.value)}
                    placeholder={t('settings.baseUrlPlaceholder')}
                  />
                </Field>
              )}

              {PROVIDERS_WITH_MODEL_DISCOVERY.has(selectedProvider.id) ? (
                <ServerModelPicker
                  providerId={selectedProvider.id}
                  baseUrl={selectedProvider.baseUrl || DEFAULT_DISCOVERY_BASE_URLS[selectedProvider.id] || ''}
                  apiKey={selectedProvider.apiKey}
                  selectedModel={settings.selectedModel}
                  onModelChange={model => onChange({ ...settings, selectedModel: model })}
                />
              ) : (
                <Field label={t('settings.modelLabel')}>
                  {models.length > 0 ? (
                    <Select
                      value={settings.selectedModel}
                      onChange={(_, data) => onChange({ ...settings, selectedModel: data.value })}
                    >
                      <option value="">{t('settings.modelSelectPlaceholder')}</option>
                      {models.map(m => (
                        <option key={m} value={m}>{m}</option>
                      ))}
                    </Select>
                  ) : (
                    <Input
                      value={settings.selectedModel}
                      onChange={(_, data) => onChange({ ...settings, selectedModel: data.value })}
                      placeholder={t('settings.modelPlaceholder')}
                    />
                  )}
                </Field>
              )}
            </>
          )}
        </div>

        <Divider />

        {/* Execution */}
        <div className={styles.section}>
          <Text weight="semibold" size={300}>{t('settings.executionSection')}</Text>

          <Field label={t('settings.autoApproveLabel')}>
            <Switch
              checked={settings.autoApprove}
              onChange={(_, data) => onChange({ ...settings, autoApprove: data.checked })}
            />
          </Field>

          <Field label={t('settings.maxRetriesLabel')}>
            <Input
              type="number"
              value={String(settings.maxRetries)}
              onChange={(_, data) => onChange({ ...settings, maxRetries: parseInt(data.value) || 3 })}
              min="1"
              max="10"
            />
          </Field>

          <Field label={t('settings.timeoutLabel')}>
            <Input
              type="number"
              value={String(settings.executionTimeout / 1000)}
              onChange={(_, data) => onChange({ ...settings, executionTimeout: (parseInt(data.value) || 30) * 1000 })}
              min="5"
              max="120"
            />
          </Field>
        </div>

        <Divider />

        {/* MCP Servers */}
        <div className={styles.section}>
          <div className={styles.row}>
            <Text weight="semibold" size={300}>{t('settings.mcpSection')}</Text>
            <Button appearance="subtle" icon={<Add24Regular />} size="small" onClick={addMcpServer}>
              {t('settings.mcpAddButton')}
            </Button>
          </div>

          {settings.mcpServers.length === 0 && (
            <Text size={200} italic>{t('settings.mcpNoServers')}</Text>
          )}

          {settings.mcpServers.map((server, i) => (
            <div key={i} className={styles.mcpServer}>
              <div className={styles.row}>
                <Input
                  value={server.name}
                  onChange={(_, data) => updateMcpServer(i, { name: data.value })}
                  placeholder={t('settings.mcpNamePlaceholder')}
                  size="small"
                  style={{ flex: 1 }}
                />
                <Switch
                  checked={server.enabled}
                  onChange={(_, data) => updateMcpServer(i, { enabled: data.checked })}
                />
                <Button
                  appearance="subtle"
                  icon={<Delete24Regular />}
                  size="small"
                  onClick={() => removeMcpServer(i)}
                />
              </div>
              <Input
                value={server.url}
                onChange={(_, data) => updateMcpServer(i, { url: data.value })}
                placeholder={t('settings.mcpUrlPlaceholder')}
                size="small"
              />
              <Select
                value={server.transport}
                onChange={(_, data) => updateMcpServer(i, { transport: data.value as McpServerConfig['transport'] })}
                size="small"
              >
                <option value="streamable-http">Streamable HTTP</option>
                <option value="sse">SSE</option>
              </Select>
            </div>
          ))}
        </div>

        <Divider />

        {/* Language */}
        <div className={styles.section}>
          <Text weight="semibold" size={300}>{t('settings.languageSection')}</Text>
          <Select
            value={locale}
            aria-label={t('settings.languageSection')}
            onChange={(_, data) => { void setLocale(data.value as LocaleId); }}
          >
            {availableLocales().map(l => (
              <option key={l.id} value={l.id}>{l.nativeName}</option>
            ))}
          </Select>
        </div>
      </div>
    </div>
  );
}

interface ServerModelPickerProps {
  providerId: string;
  baseUrl: string;
  apiKey?: string;
  selectedModel: string;
  onModelChange: (model: string) => void;
}

interface DiscoveryStatusKeys {
  checking: TranslationKey;
  connected: TranslationKey;
  unreachable: TranslationKey;
  hint: TranslationKey;
  refreshAria: TranslationKey;
}

/** Status strings differ per provider so each hint can name the right server. */
const DISCOVERY_STATUS_KEYS: Record<string, DiscoveryStatusKeys> = {
  lmstudio: {
    checking: 'settings.lmstudioStatusChecking',
    connected: 'settings.lmstudioStatusConnected',
    unreachable: 'settings.lmstudioStatusUnreachable',
    hint: 'settings.lmstudioCorsHint',
    refreshAria: 'settings.lmstudioRefreshAria',
  },
  openwebui: {
    checking: 'settings.openwebuiStatusChecking',
    connected: 'settings.openwebuiStatusConnected',
    unreachable: 'settings.openwebuiStatusUnreachable',
    hint: 'settings.openwebuiHint',
    refreshAria: 'settings.openwebuiRefreshAria',
  },
};

function ServerModelPicker({ providerId, baseUrl, apiKey, selectedModel, onModelChange }: ServerModelPickerProps) {
  const styles = useStyles();
  const { t } = useTranslation();
  const statusKeys = DISCOVERY_STATUS_KEYS[providerId] ?? DISCOVERY_STATUS_KEYS.lmstudio;
  const [status, setStatus] = useState<DiscoveredModelList['status']>('connecting');
  const [models, setModels] = useState<DiscoveredModel[]>([]);
  const abortRef = useRef<AbortController | null>(null);

  const runDiscovery = useCallback((force: boolean) => {
    abortRef.current?.abort();
    const ctrl = new AbortController();
    abortRef.current = ctrl;
    setStatus('connecting');
    discoverProviderModels(providerId, baseUrl, { signal: ctrl.signal, force, apiKey })
      .then(result => {
        if (ctrl.signal.aborted) return;
        setStatus(result.status);
        setModels(result.models);
      })
      .catch(err => {
        if (ctrl.signal.aborted) return;
        if (err instanceof DOMException && err.name === 'AbortError') return;
        setStatus('unreachable');
        setModels([]);
      });
  }, [baseUrl, apiKey, providerId]);

  // Debounced auto-discovery on mount and baseUrl change.
  useEffect(() => {
    const timer = setTimeout(() => runDiscovery(false), 500);
    return () => {
      clearTimeout(timer);
      abortRef.current?.abort();
    };
  }, [baseUrl, runDiscovery]);

  const hasModels = models.length > 0;

  return (
    <Field
      label={t('settings.modelLabel')}
      hint={
        <div className={styles.statusRow}>
          {status === 'connecting' && (
            <>
              <Spinner size="extra-tiny" />
              <Text size={200}>{t(statusKeys.checking)}</Text>
            </>
          )}
          {status === 'connected' && (
            <Badge appearance="outline" size="small" color="success">
              {t(statusKeys.connected, { count: models.length })}
            </Badge>
          )}
          {status === 'unreachable' && (
            <Tooltip content={t(statusKeys.hint)} relationship="description">
              <Badge appearance="outline" size="small" color="danger">
                {t(statusKeys.unreachable)}
              </Badge>
            </Tooltip>
          )}
          <Button
            appearance="subtle"
            size="small"
            icon={<ArrowClockwise24Regular />}
            onClick={() => runDiscovery(true)}
            aria-label={t(statusKeys.refreshAria)}
          />
        </div>
      }
    >
      {hasModels ? (
        <Select
          value={selectedModel}
          onChange={(_, data) => onModelChange(data.value)}
        >
          <option value="">{t('settings.modelSelectPlaceholder')}</option>
          {models.map(m => {
            const label = m.displayName ?? m.id;
            return (
              <option key={m.id} value={m.id}>
                {m.state === 'loaded' ? `● ${label}` : label}
              </option>
            );
          })}
        </Select>
      ) : (
        <Input
          value={selectedModel}
          onChange={(_, data) => onModelChange(data.value)}
          placeholder={t('settings.modelPlaceholder')}
        />
      )}
    </Field>
  );
}
