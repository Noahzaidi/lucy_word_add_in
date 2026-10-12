<p align="center">
  <img src="public/assets/icon-256.png" alt="AutoOffice logo" width="128" height="128" />
</p>

<h1 align="center">AutoOffice</h1>

<p align="center">AI-powered Microsoft Word + Excel + PowerPoint + Outlook add-in that writes and executes real <code>office.js</code> code on demand.</p>

## Tighten + Lucy Gateway (this fork)

This fork adds two things on top of autoOffice. Design, verification and limitations are in [NOTES.md](NOTES.md).

- **Lucy Gateway** ([gateway/](gateway/)): a small HTTPS server that holds the model API key. The add-in calls it with a bearer token, and it forwards to any OpenAI-compatible endpoint. With the "Lucy Gateway" provider selected, no model key is in the add-in.
- **Tighten** (Word): select paragraphs and click the wand button. The model proposes a tighter version of each paragraph, and you see a before/after preview. **Accept** writes it, **Reject** leaves the document untouched, and **Undo** restores the original. Tighten never runs model-written code, and headings and list items keep their style.

**You need:** Node 22+, Chrome or Edge, a free Microsoft account (Word on the web), and an OpenAI-compatible API key.

```bash
git clone <this-repo-url> autoOffice && cd autoOffice
npm install
npm run certs                          # trust https://localhost (accept the prompt)
cp gateway/.env.example gateway/.env   # fill in GATEWAY_TOKEN, UPSTREAM_API_KEY, ALLOWED_MODELS
npm run gateway                        # terminal 1: https://localhost:8787
npm run dev                            # terminal 2: https://localhost:3721
```

**Then in Word on the web**
1. Open a document, then go to **Home → Add-ins → More Add-ins → My Add-ins → Upload My Add-in** and pick `manifest.xml`.
2. In the add-in, open **Settings**:
   - Provider: **Lucy Gateway**
   - Gateway token: the value of `GATEWAY_TOKEN`
   - Model: one of `ALLOWED_MODELS`, e.g. `gpt-4.1-mini`
3. Select paragraphs and click the wand (**Tighten**) button.

**Checks**
- `npm test` and `npm run typecheck` (no Office needed).
- `curl https://localhost:8787/health` should return `{"ok":true}`. On Windows, add `--ssl-no-revoke`, because Windows' curl can't check revocation for a local dev certificate.

## What It Does

AutoOffice is a task-pane add-in you chat with. Describe what you want — for Word ("make all headings blue", "insert a 3-column table"), Excel ("put 1 through 10 in column A", "build a column chart from B2:D8"), or PowerPoint ("add a slide titled 'Q3 plan' with three bullets", "make every slide title bold") — and the agent:

1. Looks up the relevant `office.js` API docs as needed
2. Generates working code
3. Shows you the code for approval before running it
4. Executes it in a sandboxed iframe against your live Word document
5. Self-heals on errors — feeds the error back to the LLM and retries up to 3 times

## Why AutoOffice

Two things make AutoOffice different from every other AI option for Word and Excel:

**1. Private by design — open source, runs locally, works with local models.**
The whole add-in is MIT-licensed and runs in your browser as a normal Office task pane. Your API key is stored on your machine, and traffic goes only to the provider you pick — Anthropic, OpenAI, Azure, or any OpenAI-compatible endpoint. Point it at Ollama or LM Studio and your document never leaves the machine. No Microsoft account requirement, no Copilot license, no third-party server in the middle. Audit the code, fork it, self-host the build — nothing is hidden.

**2. Pure `office.js` + MCP — minimal in code, maximal in capability.**
There is no curated wrapper API. The agent has exactly two built-in tools: look up the relevant `office.js` API docs, and run real code against the live document. That means anything Word or Excel can natively do is reachable on day one — formatting, tables, charts, ranges, comments, content controls, all of it — without us having to "support" each operation one at a time. Plug in MCP servers (Linear, Notion, internal APIs) and the same tiny surface area extends to anything outside Office too. Tiny core, unbounded reach.

## Comparison: Word + Excel + PowerPoint AI Add-ins

| | **AutoOffice** | **Microsoft Copilot** | **Claude for Word** | **Word GPT Plus** |
|---|:---:|:---:|:---:|:---:|
| **— Privacy & openness —** | | | | |
| Open source | ✅ MIT | ❌ | ❌ | ✅ MIT |
| Self-hostable | ✅ | ❌ | ❌ | ✅ |
| Local model support | ✅ Ollama, any OpenAI-compatible endpoint | ❌ | ❌ | ✅ Ollama |
| BYO API key (key stays local) | ✅ | ❌ | ❌ | ✅ |
| Cost | Free + your provider usage | M365 Copilot license | Claude paid plan | Free + your provider usage |
| Provider choice | Anthropic, OpenAI, any OpenAI-compatible, Ollama | Microsoft-hosted only | Claude only | OpenAI, Azure, Gemini, Ollama |
| **— Power (small core, big reach) —** | | | | |
| Executes real `office.js` | ✅ | ❌ | ❌ | ⚠️ partial |
| Code preview before run | ✅ | ❌ | ❌ | ❌ |
| MCP support | ✅ HTTP / SSE | ✅ via Copilot Studio | ❌ | ❌ |
| Reaches full Word + Excel API surface | ✅ via real code | ⚠️ curated actions | ⚠️ curated actions | ⚠️ curated actions |
| **— AutoOffice trade-offs —** | | | | |
| Native tracked-changes UX | ❌ | ❌ | ✅ | ❌ |
| Cross-app context (Outlook, Teams, PPT…) | ❌ Word + Excel + PowerPoint, single-doc | ✅ all M365 | ✅ Word + Excel + PPT | ❌ |

## Install

Prebuilt Windows installer — no Node, no source build, no dev cert.

1. Download **AutoOffice-Setup.exe** from the [latest release](https://github.com/matthamiltonnz/autoOffice/releases/latest) and run it.
2. If Windows shows **"Windows protected your PC"**, click **More info → Run anyway** (the installer is unsigned).
3. Restart Word (or Excel / PowerPoint).
4. In Word: **Home → Add-ins → Advanced → Shared Folder**, pick **AutoOffice**, and click **Add**.

The installer just registers an Office add-in manifest. The task-pane assets are served from this fork's GitHub Pages site at `https://matthamiltonnz.github.io/autoOffice/`, so it runs entirely on its own hosting and loads no code from upstream — see the self-hosting notes below.

### Self-hosting

This fork is self-hosted on GitHub Pages, so it serves no code from upstream:

- **Task-pane assets:** `https://matthamiltonnz.github.io/autoOffice/`
- **Repo:** `https://github.com/matthamiltonnz/autoOffice`

`deploy.yml` keeps `VITE_BASE: /autoOffice/` because the repo is named `autoOffice` —
that stays correct as long as the repo name doesn't change.

To set this up from scratch (or on a fresh fork):

1. **Create the repo**, e.g. `matthamiltonnz/autoOffice`. It must be **public** —
   GitHub Pages on the free plan does not serve private repositories.
2. In the repo: **Settings → Pages → Source: GitHub Actions**.
3. If the repo is **not** named `autoOffice`, edit `.github/workflows/deploy.yml` and change
   `VITE_BASE: /autoOffice/` to `/<your-repo-name>/`.
4. Push to `master` (or run **Deploy to GitHub Pages** manually). Assets are served at
   `https://matthamiltonnz.github.io/autoOffice/`.
5. Confirm the URLs in **both** production manifests match that origin — `manifest.production.xml`
   (Word/Excel/PowerPoint) and `manifest.outlook.production.xml` (Outlook). Each carries its own
   `<Id>` GUID; keep them distinct from upstream's so both add-ins can coexist on one machine.
6. Register the add-in for users, either by:
   - rebuilding the installer (`installer/setup.iss` via Inno Setup) so it ships the edited
     `manifest.production.xml`, or
   - sideloading the manifest directly: **Insert → Add-ins → Upload My Add-in** (desktop),
     or copying it into Outlook's `wef` folder on macOS:
     `~/Library/Containers/com.microsoft.Outlook/Data/Documents/wef/`.

Because Pages serves static files, **nothing needs to stay running** — no server, no container,
and no dev machine that has to be switched on.

## Development

### Prerequisites

- Node.js 18+
- Microsoft 365 (Word, Excel, or PowerPoint — Web or Desktop)
- An API key for Anthropic, OpenAI, or any OpenAI-compatible provider

### Install

```bash
npm install
```

### Dev certs (required for Office sideloading)

```bash
npm run certs
```

This installs a self-signed localhost certificate so Word will trust the add-in URL.

### Run + sideload (with debugger)

```bash
npm run start
```

Starts the dev server and sideloads the add-in into Word with the debugger attached. The server runs at **https://localhost:3721**.

### Run + sideload (no debugger — recommended for regular dev)

```bash
npm run sideload
```

Same as `start` but without attaching the debugger. Faster startup, targets desktop Word directly. Use this for day-to-day testing when you don't need breakpoints.

### Run + sideload Excel

Same scripts but targeting Excel:

```bash
npm run start:excel       # debugger
npm run sideload:excel    # no debugger
```

### Run + sideload PowerPoint

Same scripts but targeting PowerPoint:

```bash
npm run start:powerpoint       # debugger
npm run sideload:powerpoint    # no debugger
```

### Run + sideload Outlook

Outlook needs its own manifest — `manifest.outlook.xml` — because an add-in manifest can contain only one `<VersionOverrides>` element, and the Mailbox host uses the `mailappversionoverrides` namespace while Word/Excel/PowerPoint use `taskpaneappversionoverrides`. Both manifests point at the same web app.

**macOS** — sideload into Outlook's `wef` folder, then restart Outlook:

```bash
npm run sideload:outlook:mac
```

**Windows** — Office debugging tooling:

```bash
npm run start:outlook
```

Then in Outlook: **Home → Add-ins → AutoOffice**. In Outlook on the web, open **<https://aka.ms/olksideload>** — the *Add-Ins for Outlook* dialog opens directly — then **My add-ins → Custom Addins → Add a custom add-in → Add from File**. `Add from URL` is no longer supported for Outlook. On Outlook for Mac 16.85+ the ribbon **Get Add-ins** button opens the Microsoft Marketplace in the browser instead of this dialog, so use the URL rather than hunting for the button. Upload `manifest.outlook.xml` (or `manifest.outlook.production.xml` for a hosted build — edit its URLs to your Pages domain first, exactly as described above).

Surfaces: the manifest declares both `MessageReadCommandSurface` and `MessageComposeCommandSurface`, each with a matching form in `FormSettings` (`ItemRead` and `ItemEdit`). Both are needed — a compose surface whose `ItemEdit` form is missing opens to a blank pane with no error. Capabilities and limits: the manifest requests **Mailbox 1.3** and **ReadWriteItem**, so the agent can read the open message or meeting (subject, sender, recipients, body, attachment metadata) and write drafts (body, subject, recipients, reply/forward forms, attachments). Reading other messages or folders would need `ReadWriteMailbox` plus EWS/REST, which is deliberately out of scope.


### Run dev server only

```bash
npm run dev
```

Starts only the Vite dev server — no sideloading. Useful if you're working on the UI and sideloading separately.

### Sideload manually

If the dev server is already running:

**Insert → Add-ins → Upload My Add-in** → pick `manifest.xml`.

### Stop

```bash
npm run stop
```

### Configure

Open the add-in task pane and click the settings gear:

- **Provider:** Anthropic, OpenAI, or any OpenAI-compatible endpoint (Ollama, LM Studio, Open WebUI, etc.)
- **Base URL:** for self-hosted providers — `http://localhost:1234/v1` for LM Studio, `http://localhost:8080/api` for Open WebUI
- **API Key:** stored locally, never sent anywhere except directly to the provider
- **Model:** e.g. `claude-opus-4-7`, `gpt-4o` — for LM Studio and Open WebUI the list is fetched from the server, so you pick it from a dropdown
- **Auto-approve:** skip the approve step and run code immediately
- **MCP Servers:** add HTTP/SSE MCP servers to extend the agent with external tools

## Project Structure

```
src/taskpane/
├── index.tsx              — Entry point, Office.onReady
├── App.tsx                — Root component, state management
├── agent/
│   ├── orchestrator.ts    — Agent loop: streamText + tool routing + self-healing
│   ├── tools.ts           — Built-in tool definitions
│   └── providers.ts       — Provider factory (Anthropic, OpenAI, compatible)
├── components/
│   ├── ChatPanel.tsx      — Message list + input
│   ├── CodeBlock.tsx      — Syntax-highlighted code with approve/reject
│   ├── MessageBubble.tsx  — Individual message
│   ├── ToolActivity.tsx   — Inline tool call indicators
│   └── SettingsPanel.tsx  — Provider, API keys, MCP, auto-approve
├── executor/
│   ├── sandbox.ts         — Iframe lifecycle + postMessage bridge
│   └── iframe.html        — Sandbox page (loads office.js, receives execute messages)
├── skills/                — office.js API doc chunks (markdown, one per domain)
│   ├── index.ts           — Registry + lookup function
│   ├── formatting.md
│   ├── tables.md
│   ├── context-sync.md    — Critical: load()/sync() batching model
│   └── ...
├── mcp/
│   └── client.ts          — MCP client via @ai-sdk/mcp
└── store/
    └── settings.ts        — Persist settings (roamingSettings in Office, localStorage in dev)
```

## Built-in Agent Tools

| Tool | What it does |
|------|-------------|
| `lookup_skill(name)` | Fetches `office.js` API docs for a domain (formatting, tables, ranges, etc.) |
| `execute_code(code)` | Submits generated code to the sandboxed iframe for execution |

To read document state the agent writes `execute_code` that loads and returns the needed properties directly — no separate tool required.

MCP server tools are surfaced alongside these automatically.

## Self-Healing

When code execution fails, the error is fed back to the LLM with the instruction to fix it. Each retry appears as a visible message in chat. After 3 failures the agent gives up and shows the final error.

## Settings

| Setting | Default |
|---------|---------|
| AI Provider | (none — select in settings) |
| Model | (provider-dependent) |
| Auto-approve | Off |
| Max retries | 3 |
| Execution timeout | 30 seconds |
| MCP Servers | Empty |

Settings are persisted via `Office.context.roamingSettings` when running inside Office, and `localStorage` during development. Provider, API key, MCP server, and other settings are shared across Word, Excel, and PowerPoint by design — there is one logical add-in per install.

## Build

```bash
npm run build
```

Output goes to `dist/`. Deploy the `dist/` folder to any HTTPS host and update the URLs in `manifest.xml`.

## Notes

- **Browser-only MCP:** The add-in runs entirely client-side. Only HTTP/SSE MCP transports work — no stdio. Local MCP servers need to expose an HTTP endpoint.
- **CORS:** Direct browser-to-API calls work with Anthropic and OpenAI. If you hit CORS issues with a provider, you'll need a lightweight proxy.
- **iframe context:** The sandbox iframe loads its own `office.js` instance. This follows the same pattern as Microsoft's Script Lab.

## Roadmap

### Chat History
Conversations are persisted to `localStorage` per-device and restored on reload. You can keep multiple named conversations and switch between them from the history panel; titles are auto-generated by the configured model after the first turn.

Planned work:
- **Export** — copy conversation (including generated code and results) as markdown or plain text

## Adding a Language

AutoOffice's UI and AI agent both run in the user's language.

To add a new language (example: French):

1. **Register the locale** in `src/taskpane/i18n/registry.ts`:

   ```ts
   export const LOCALES = {
     en: { name: 'English', nativeName: 'English',  direction: 'ltr', fallback: null },
     he: { name: 'Hebrew',  nativeName: 'עברית',    direction: 'rtl', fallback: 'en' },
     fr: { name: 'French',  nativeName: 'Français', direction: 'ltr', fallback: 'en' },
   } as const satisfies Record<string, LocaleMeta>;
   ```

2. **Translate the strings.** Copy `src/taskpane/i18n/locales/en.json` to `src/taskpane/i18n/locales/fr.json` and translate the values. Keep the keys identical and the `{{name}}` placeholders intact.

3. **Verify coverage:**

   ```bash
   npm run check:i18n
   ```

   You should see `✓ fr: N keys, complete`. Missing keys fail the build; extra keys are warnings.

That's it — the locale appears in the Settings → Language dropdown automatically, and the AI agent will reply in the new language for users who pick it.

## Tech Stack

- **Framework:** React 19 + TypeScript
- **Build:** Vite with HTTPS (required for Office sideloading)
- **UI:** Fluent UI (`@fluentui/react-components`)
- **AI:** Vercel AI SDK (`ai`, `@ai-sdk/anthropic`, `@ai-sdk/openai`)
- **MCP:** `@ai-sdk/mcp` for external tool servers
- **Code highlighting:** Shiki
- **Schemas:** Zod

## Architecture

```
Task Pane (React)
├── Chat UI (Fluent UI)          ← user input, message history, tool activity
├── Code Preview Block           ← syntax-highlighted code with Approve / Reject
└── Agent Orchestrator
    ├── streamText (AI SDK)      ← multi-provider LLM client
    ├── Skill Registry           ← office.js API docs fetched on demand
    ├── MCP Client               ← external tool servers (HTTP only)
    └── postMessage bridge
            └── Sandboxed iframe ← executes generated code against live document
```

The same task pane runs in Word, Excel, and PowerPoint; `HostContext` is resolved at startup and gates the skill registry, sandbox wrapping, and system prompt per host.
