// @vitest-environment node
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { createApp } from './app.ts';

const GATEWAY_TOKEN = 'test-gateway-token';
const UPSTREAM_KEY = 'sk-upstream-secret';

const config = {
  gatewayToken: GATEWAY_TOKEN,
  upstreamBaseUrl: 'https://upstream.example/v1',
  upstreamApiKey: UPSTREAM_KEY,
  allowedModels: ['gpt-4.1-mini'],
  allowedOrigins: ['https://localhost:3721'],
};

const fetchMock = vi.fn();

function postChat(body: unknown, authorization = `Bearer ${GATEWAY_TOKEN}`) {
  return createApp(config).request('/v1/chat/completions', {
    method: 'POST',
    headers: { Authorization: authorization, 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
}

const allowedRequest = { model: 'gpt-4.1-mini', messages: [{ role: 'user', content: 'hi' }] };

beforeEach(() => {
  fetchMock.mockReset();
  vi.stubGlobal('fetch', fetchMock);
  vi.spyOn(console, 'log').mockImplementation(() => {});
});

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe('gateway', () => {
  it('answers the browser preflight from the add-in origin', async () => {
    const res = await createApp(config).request('/v1/chat/completions', {
      method: 'OPTIONS',
      headers: {
        Origin: 'https://localhost:3721',
        'Access-Control-Request-Method': 'POST',
        'Access-Control-Request-Headers': 'authorization,content-type',
      },
    });

    expect(res.status).toBe(204);
    expect(res.headers.get('Access-Control-Allow-Origin')).toBe('https://localhost:3721');
  });

  it('rejects a request without a token', async () => {
    const res = await createApp(config).request('/v1/chat/completions', {
      method: 'POST',
      body: JSON.stringify(allowedRequest),
    });

    expect(res.status).toBe(401);
    expect(fetchMock).not.toHaveBeenCalled();
  });

  it('keeps CORS headers on the 401, so the add-in can read the error', async () => {
    const res = await createApp(config).request('/v1/chat/completions', {
      method: 'POST',
      headers: { Origin: 'https://localhost:3721' },
      body: JSON.stringify(allowedRequest),
    });

    expect(res.status).toBe(401);
    expect(res.headers.get('Access-Control-Allow-Origin')).toBe('https://localhost:3721');
  });

  it('rejects a wrong token', async () => {
    const res = await postChat(allowedRequest, 'Bearer some-other-token');

    expect(res.status).toBe(401);
    expect(fetchMock).not.toHaveBeenCalled();
  });

  it('rejects a model that is not on the allowlist', async () => {
    const res = await postChat({ ...allowedRequest, model: 'gpt-expensive' });

    expect(res.status).toBe(400);
    expect(fetchMock).not.toHaveBeenCalled();
  });

  it('forwards to the upstream with the upstream key, never the gateway token', async () => {
    fetchMock.mockResolvedValue(
      new Response(JSON.stringify({ id: 'reply-1' }), { status: 200, headers: { 'Content-Type': 'application/json' } }),
    );

    const res = await postChat(allowedRequest);

    expect(res.status).toBe(200);
    expect(await res.json()).toEqual({ id: 'reply-1' });
    const [url, init] = fetchMock.mock.calls[0];
    expect(url).toBe('https://upstream.example/v1/chat/completions');
    expect(init.headers.Authorization).toBe(`Bearer ${UPSTREAM_KEY}`);
    expect(JSON.stringify(init.headers)).not.toContain(GATEWAY_TOKEN);
    expect(JSON.parse(init.body)).toEqual(allowedRequest);
  });

  it('streams a server-sent-events reply through unchanged', async () => {
    const events = 'data: {"choices":[{"delta":{"content":"Hi"}}]}\n\ndata: [DONE]\n\n';
    fetchMock.mockResolvedValue(new Response(events, { status: 200, headers: { 'Content-Type': 'text/event-stream' } }));

    const res = await postChat({ ...allowedRequest, stream: true });

    expect(res.headers.get('Content-Type')).toBe('text/event-stream');
    expect(await res.text()).toBe(events);
  });

  it('rejects a body that is not JSON', async () => {
    const res = await createApp(config).request('/v1/chat/completions', {
      method: 'POST',
      headers: { Authorization: `Bearer ${GATEWAY_TOKEN}`, 'Content-Type': 'application/json' },
      body: 'not json',
    });

    expect(res.status).toBe(400);
    expect(fetchMock).not.toHaveBeenCalled();
  });

  it('hides the upstream auth error behind a generic 502', async () => {
    fetchMock.mockResolvedValue(
      new Response(JSON.stringify({ error: { message: 'Incorrect API key provided: sk-upst****cret' } }), { status: 401 }),
    );

    const res = await postChat(allowedRequest);

    expect(res.status).toBe(502);
    expect(await res.text()).not.toContain('sk-');
  });

  it('returns 502 when the upstream is unreachable', async () => {
    fetchMock.mockRejectedValue(new TypeError('fetch failed'));

    const res = await postChat(allowedRequest);

    expect(res.status).toBe(502);
  });
});
