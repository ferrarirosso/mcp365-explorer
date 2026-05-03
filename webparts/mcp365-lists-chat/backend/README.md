# mcp365-lists-chat — backend

Azure Functions LLM proxy for the C1 webpart. Routes browser `chat/completions` calls to Azure OpenAI with managed identity, CORS locked to your SharePoint tenant, and per-caller rate limiting.

## What it does

One endpoint: `POST /api/chat/completions` — a thin passthrough to Azure OpenAI. The browser sends an OpenAI-compatible request body (messages + tool definitions); the proxy forwards it to your Azure OpenAI deployment. The tool-calling loop runs in the browser, not here.

## Why a proxy at all?

Three reasons the browser can't call Azure OpenAI directly:
1. **API key secrecy** — browser bundles are public
2. **CORS** — AOAI doesn't allow cross-origin browser requests
3. **Rate limiting** — per-caller quotas need server-side enforcement

## Prerequisites

- Azure CLI: https://learn.microsoft.com/cli/azure/install-azure-cli
- Azure Functions Core Tools: https://learn.microsoft.com/azure/azure-functions/functions-run-local
- Node.js 22+
- An Azure subscription with permission to create Cognitive Services resources

## Deploy

```bash
npm install
npm run deploy
```

Interactive prompts. Takes ~5 minutes. Creates:
- Resource group
- Azure AI Services resource (kind: `AIServices`)
- `gpt-5-mini` model deployment
- Storage account
- Function App (Node 22, Linux, Consumption)
- System-assigned managed identity
- Role assignment (Function → AI Services)

At the end it prints the proxy URL and the Backend API resource. The deployer auto-wires both into the webpart's `serve.json`. No function key in the browser path — Easy Auth gates the proxy via the SPFx-acquired Entra token.

## Teardown

```bash
npm run teardown
```

Type the resource group name to confirm. Deletes everything. Takes 2–5 minutes in the background.

## Local development

```bash
npm run setup     # generate local.settings.json interactively
npm start         # func start on :7071
curl http://localhost:7071/api/health
```

For local dev you can use either:
- **Managed identity** (run `az login` first — picks up your credentials via `DefaultAzureCredential`)
- **API key** (paste it during `npm run setup`)

## Configuration

All configured via environment variables, set by `deploy.mjs` or `setup-local.mjs`:

| Env var | Purpose |
|---|---|
| `AZURE_OPENAI_ENDPOINT` | AI Services endpoint |
| `AZURE_OPENAI_API_VERSION` | Defaults to `2025-01-01-preview` |
| `AZURE_OPENAI_DEPLOYMENT` | Name of the model deployment |
| `AZURE_OPENAI_API_KEY` | Optional (dev). If absent, uses managed identity. |
| `AZURE_OPENAI_REASONING_EFFORT` | `low` / `medium` / `high`. Default `low`. |
| `REQUESTS_PER_MINUTE` | Per-caller rate limit (default 30) |
| `REQUESTS_PER_DAY` | Per-caller daily limit (default 1000) |
| `ALLOWED_ORIGIN` | CORS origin (your SharePoint tenant URL) |
| `ALLOW_PERMISSIVE_LOCAL_CORS` | `true` in dev, `false` in prod |

## Routes

| Method | Path | Purpose |
|---|---|---|
| `GET`  | `/api/health` | Public readiness — `{ status: "ok", time }`. Returns 401 once Easy Auth is enabled. |
| `POST` | `/api/chat/completions` | OpenAI-compatible chat completions passthrough |
| `OPTIONS` | (both) | CORS preflight |

All non-preflight routes require an `Authorization: Bearer <token>` header — a token from `AadTokenProvider.getToken(backendApiResource)`. Easy Auth (audience-pinned) validates it; no function key.

## What's NOT in C1

Kept out on purpose to keep the post tight:
- Easy Auth / Entra ID backend API registration → comes back in C6 (approval gates / productionization)
- Per-user identity via Easy Auth pass-through → same
- Streaming responses → maybe a refinement post
- Persistent rate limiting (CosmosDB/Redis) → production concern, not demo concern

Current rate limiting is per caller IP (from `X-Forwarded-For`), in-memory — resets when the function restarts. Fine for demos, not for production.
