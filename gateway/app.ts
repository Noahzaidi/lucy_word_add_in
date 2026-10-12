import { Hono } from 'hono';
import { cors } from 'hono/cors';
import { bearerAuth } from 'hono/bearer-auth';
import { bodyLimit } from 'hono/body-limit';

export interface GatewayConfig {
  gatewayToken: string;
  upstreamBaseUrl: string;
  upstreamApiKey: string;
  allowedModels: string[];
  allowedOrigins: string[];
}

const MAX_BODY_BYTES = 1_000_000;

/** OpenAI's error shape, so the add-in shows our message like any provider error. */
function openAiError(message: string) {
  return { error: { message } };
}

export function createApp(config: GatewayConfig): Hono {
  const app = new Hono();

  // CORS runs first: it answers the browser's preflight itself and keeps CORS
  // headers on error responses so the add-in can read them. It is not access
  // control — any non-browser client ignores it. The bearer token is the control.
  app.use('*', cors({ origin: config.allowedOrigins }));

  app.get('/health', (c) => c.json({ ok: true }));

  app.use('/v1/*', bearerAuth({ token: config.gatewayToken }));
  app.use('/v1/*', bodyLimit({ maxSize: MAX_BODY_BYTES }));

  app.post('/v1/chat/completions', async (c) => {
    let body: { model?: unknown } | null;
    try {
      body = await c.req.json();
    } catch {
      return c.json(openAiError('Request body must be JSON.'), 400);
    }

    const model = body?.model;
    if (typeof model !== 'string' || !config.allowedModels.includes(model)) {
      return c.json(openAiError(`Model not allowed. Allowed: ${config.allowedModels.join(', ')}`), 400);
    }

    // Fresh headers only: the add-in's Authorization header carries the gateway
    // token, which must never reach the upstream provider.
    const startedAt = Date.now();
    let upstream: Response;
    try {
      upstream = await fetch(`${config.upstreamBaseUrl}/chat/completions`, {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${config.upstreamApiKey}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(body),
        // Stops the upstream call (and its token spend) if the add-in disconnects.
        signal: c.req.raw.signal,
      });
    } catch (err) {
      const reason = c.req.raw.signal.aborted
        ? 'cancelled by the add-in'
        : `upstream unreachable: ${err instanceof Error ? err.message : String(err)}`;
      console.log(`${model} -> ${reason}`);
      return c.json(openAiError('Upstream unreachable.'), 502);
    }
    // Log metadata only, never prompt content.
    console.log(`${model} -> ${upstream.status} (${Date.now() - startedAt} ms)`);

    // The upstream's own message would blame the add-in's token and can echo
    // part of the real key, so replace it with a generic one.
    if (upstream.status === 401 || upstream.status === 403) {
      await upstream.body?.cancel();
      return c.json(openAiError("Upstream rejected the gateway's credentials."), 502);
    }

    // Pass the body through as-is: a streamed (SSE) reply streams to the add-in.
    // Only Content-Type is copied; Node's fetch has already decoded the body, so
    // the upstream's content-encoding and content-length would be wrong here.
    return new Response(upstream.body, {
      status: upstream.status,
      headers: { 'Content-Type': upstream.headers.get('Content-Type') ?? 'application/json' },
    });
  });

  return app;
}
