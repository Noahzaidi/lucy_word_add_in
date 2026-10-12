import fs from 'node:fs';
import https from 'node:https';
import os from 'node:os';
import path from 'node:path';
import { serve } from '@hono/node-server';
import { createApp } from './app.ts';

// The characters Hono's bearer-auth accepts; any other token gets a 400.
const TOKEN_PATTERN = /^[A-Za-z0-9._~+/-]+=*$/;
// The token is the gateway's only access control, so refuse short, guessable ones.
const MIN_TOKEN_LENGTH = 32;

function requireEnv(name: string): string {
  const value = process.env[name]?.trim();
  if (!value) {
    throw new Error(`${name} is missing from gateway/.env (see gateway/.env.example).`);
  }
  return value;
}

function splitList(value: string): string[] {
  return value.split(',').map((item) => item.trim()).filter(Boolean);
}

/** Header values must be printable ASCII; a stray copy-paste character breaks every request. */
function findInvalidCharacter(value: string): number {
  for (let i = 0; i < value.length; i++) {
    const code = value.charCodeAt(i);
    if (code < 0x21 || code > 0x7e) return i;
  }
  return -1;
}

function loadConfig() {
  const gatewayToken = requireEnv('GATEWAY_TOKEN');
  if (gatewayToken.length < MIN_TOKEN_LENGTH || !TOKEN_PATTERN.test(gatewayToken)) {
    throw new Error(
      `GATEWAY_TOKEN must be at least ${MIN_TOKEN_LENGTH} characters of letters, digits and . _ ~ + / - (generate one with the command in .env.example).`,
    );
  }

  const upstreamApiKey = requireEnv('UPSTREAM_API_KEY');
  const invalidAt = findInvalidCharacter(upstreamApiKey);
  if (invalidAt !== -1) {
    throw new Error(`UPSTREAM_API_KEY has an invalid character at position ${invalidAt + 1}, usually a copy-paste artifact. Re-copy the key.`);
  }

  return {
    gatewayToken,
    upstreamApiKey,
    upstreamBaseUrl: requireEnv('UPSTREAM_BASE_URL').replace(/\/+$/, ''),
    allowedModels: splitList(requireEnv('ALLOWED_MODELS')),
    allowedOrigins: splitList(process.env.ALLOWED_ORIGINS ?? 'https://localhost:3721'),
    port: Number(process.env.PORT ?? 8787),
  };
}

/** The same localhost certificate the add-in's dev server uses (installed by `npm run certs`). */
function loadDevCertificate() {
  const certDir = path.join(os.homedir(), '.office-addin-dev-certs');
  const certFile = path.join(certDir, 'localhost.crt');
  const keyFile = path.join(certDir, 'localhost.key');
  if (!fs.existsSync(certFile) || !fs.existsSync(keyFile)) {
    throw new Error(`Dev certificate not found in ${certDir}. Run "npm run certs" first.`);
  }
  return { cert: fs.readFileSync(certFile), key: fs.readFileSync(keyFile) };
}

try {
  const config = loadConfig();
  serve(
    {
      fetch: createApp(config).fetch,
      port: config.port,
      hostname: 'localhost',
      createServer: https.createServer,
      serverOptions: loadDevCertificate(),
    },
    (info) => {
      console.log(`Gateway listening on https://localhost:${info.port}`);
      console.log(`Forwarding to ${config.upstreamBaseUrl} for models: ${config.allowedModels.join(', ')}`);
    },
  );
} catch (err) {
  console.error(`Gateway not started: ${(err as Error).message}`);
  process.exit(1);
}
