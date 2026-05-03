# mcp365-lists-chat

Chat interface over the SharePoint Lists MCP server. The LLM picks tools from `tools/list` and fills arguments; the webpart executes them via MCP and feeds results back.

Part of the [MCP365 Explorer](https://github.com/ferrarirosso/mcp365-explorer) series — post C1 in the agentic track.

## Architecture

```
Browser (SPFx webpart)
  │
  ├──→ @ferrarirosso/mcp-browser-client ──→ Work IQ SharePoint Lists MCP
  │    (via AadTokenProvider, direct from the browser)
  │
  └──→ backend/ (Azure Functions proxy) ──→ Microsoft Foundry / Azure OpenAI
       (hides API key, enforces CORS, per-user rate limiting)

Tool-calling loop lives in the browser:
  user msg → proxy (LLM + tools/list) → tool_calls
    → McpBrowserClient.callTool() → proxy (LLM + tool result)
    → final answer → render
```

## Run locally

```bash
cd webparts/mcp365-lists-chat
npm install
npm run deploy         # provisions the proxy + auto-wires serve.json
npm start              # workbench opens with property pane pre-filled
```

The deployer ([`spfx-foundry-deploy`](https://github.com/ferrarirosso/spfx-foundry-deploy),
consumed via `npx -y github:ferrarirosso/spfx-foundry-deploy`) handles every step:
subscription pick, region pick (filtered to gpt-5-mini-friendly Foundry regions),
naming, SharePoint origin auto-suggested from the `az login` UPN domain, AI Services +
model + storage + Function App, Backend API app + scope + SPFx grant, managed
identity to Foundry, Easy Auth requiring an Entra token, platform hardening
(HTTPS-only, TLS 1.2, FTP off), Application Insights, code deploy, app settings,
health check. After it finishes you have a `serve.json` already wired with
`backendUrl` and `backendApiResource` — no copy/paste, no shared secrets.

## Property pane

| Field | How it's filled |
|---|---|
| Environment ID | Manual — Power Platform environment GUID (see `pac admin list`) |
| Backend URL | Auto-wired by `npm run deploy` |
| Backend API resource | Auto-wired by `npm run deploy` (e.g. `api://<guid>`) |

The auth chain is `AadTokenProvider.getToken(backendApiResource)` → `Authorization: Bearer …` → Easy Auth (audience-pinned) → Function code → managed identity → Foundry. No function key in the browser bundle.

## Other commands

```bash
npm run deploy:dry-run # walk through the prompts, see the plan, no Azure calls
npm run setup          # re-wires serve.json from .deploy-output.json (no Azure changes)
npm run teardown       # deletes the resource group; zero lingering infra
```

`npm run deploy:dry-run` is its own script — `npm run deploy --dry-run` doesn't work (npm consumes the flag).

For backend-side local development (`func start` against the proxy locally):

```bash
cd backend
npm run setup-local    # generates backend/local.settings.json
npm start              # func start
```
