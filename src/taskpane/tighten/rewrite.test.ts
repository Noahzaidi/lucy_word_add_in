import { describe, it, expect, vi, beforeEach } from 'vitest';
import { NoObjectGeneratedError } from 'ai';

const mocks = vi.hoisted(() => ({
  generateText: vi.fn(),
  createModel: vi.fn(),
}));

vi.mock('ai', async () => {
  const actual = await vi.importActual<typeof import('ai')>('ai');
  return { ...actual, generateText: mocks.generateText };
});
vi.mock('../agent/providers.ts', () => ({ createModel: mocks.createModel }));

import { requestRewrites } from './rewrite.ts';
import { translationService } from '../i18n/index.ts';
import type { AppSettings } from '../store/settings.ts';

const settings: AppSettings = {
  selectedProviderId: 'lucy-gateway',
  selectedModel: 'gpt-4.1-mini',
  providers: [{ id: 'lucy-gateway', name: 'Lucy Gateway', apiKey: 'token', baseUrl: 'https://localhost:8787/v1' }],
  autoApprove: false,
  mcpServers: [],
  maxRetries: 3,
  executionTimeout: 30000,
};

const paragraphs = [{ id: 0, text: 'In order to start, we need to begin.' }];

describe('requestRewrites', () => {
  beforeEach(() => {
    mocks.generateText.mockReset();
    mocks.createModel.mockReset();
    mocks.createModel.mockReturnValue('FAKE_MODEL');
  });

  it('sends the paragraphs as JSON data with a structured-output schema and no retries', async () => {
    const output = { paragraphs: [{ id: 0, text: 'To start, begin.' }] };
    mocks.generateText.mockResolvedValue({ finishReason: 'stop', output });

    const result = await requestRewrites(paragraphs, settings);

    expect(result).toEqual({ ok: true, output });
    const call = mocks.generateText.mock.calls[0][0];
    expect(JSON.parse(call.prompt)).toEqual({ paragraphs });
    expect(call.output).toBeDefined();
    expect(call.maxRetries).toBe(0);
  });

  it('reports a cut-off reply instead of using it', async () => {
    mocks.generateText.mockResolvedValue({ finishReason: 'length', output: undefined });

    const result = await requestRewrites(paragraphs, settings);

    expect(result.ok).toBe(false);
  });

  it('reports a reply that does not match the schema', async () => {
    mocks.generateText.mockRejectedValue(
      new NoObjectGeneratedError({
        message: 'response did not match schema',
        text: '{"rewrites":[]}',
        response: { id: 'r', timestamp: new Date(0), modelId: 'gpt-4.1-mini' },
        usage: { inputTokens: 1, outputTokens: 1, totalTokens: 2 } as never,
        finishReason: 'stop',
      }),
    );

    const result = await requestRewrites(paragraphs, settings);

    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.title).toBe(translationService.t('tighten.badFormatTitle'));
  });

  it('formats other errors, e.g. the gateway rejecting the token', async () => {
    mocks.generateText.mockRejectedValue(new Error('Unauthorized'));

    const result = await requestRewrites(paragraphs, settings);

    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.detail).toBe('Unauthorized');
  });
});
