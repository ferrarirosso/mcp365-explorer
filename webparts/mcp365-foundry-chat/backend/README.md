# mcp365-foundry-chat — backend

Azure Functions proxy for the showcase webpart. Routes browser `chat/completions` calls to Microsoft Foundry with managed identity, CORS locked to your SharePoint tenant, Easy Auth in pass-through mode, and per-caller rate limiting.

The showcase webpart's purpose is to prove this auth chain works end-to-end without any MCP layered on top.

## What it does

One endpoint: `POST /api/chat/completions` — a thin passthrough to Foundry. The browser sends an OpenAI-compatible request body (messages); the proxy forwards it to your Foundry deployment.

## Deploy / teardown

Run from the **webpart** directory, not from `backend/`:

```bash
cd webparts/mcp365-foundry-chat
npm run deploy        # provisions everything + auto-wires serve.json
npm run teardown      # deletes the resource group
```

See [`spfx-foundry-deploy`](https://github.com/ferrarirosso/spfx-foundry-deploy) for the per-step breakdown of what `npm run deploy` provisions.

## Local development (`func start`)

```bash
cd backend
npm run setup-local   # generate local.settings.json interactively
npm start             # func start on :7071
curl http://localhost:7071/api/health
```

Auth: managed identity (run `az login` first) or paste an API key during `setup-local`.

## Configuration (env vars)

| Env var | Purpose |
|---|---|
| `AZURE_OPENAI_ENDPOINT` | Foundry endpoint |
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

All non-preflight routes require:
- `Authorization: Bearer <token>` — a token from `AadTokenProvider.getToken(backendApiResource)`. Validated by Easy Auth (audience pinned to `api://<backend-app-id>`); the function code defensively rejects requests with no `x-ms-client-principal` header even though Easy Auth should have stopped them already.

No `x-functions-key`, no shared secret. Function keys in a public SPFx bundle would be a shared secret; the Entra bearer token is the only auth gate.
