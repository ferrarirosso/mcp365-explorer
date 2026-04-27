# mcp365-foundry-chat

Minimal chat against a protected, Foundry-backed Function App proxy. Proves the deployment + SPFx auth chain works end-to-end. **No MCP, no tool-calling.**

If this webpart can chat, every other webpart that builds on the same deployer will also work.

## Architecture

```
Browser (SPFx webpart)
  │  Authorization: Bearer <token from AadTokenProvider for api://<backend-app>>
  ↓
Azure Functions proxy
  │  Easy Auth: requireAuthentication=true, audience pinned to the Backend API
  │  Function code: Easy Auth principal extracted from x-ms-client-principal-id
  ↓  Managed identity → Cognitive Services OpenAI User
Microsoft Foundry
```

One round-trip per send: `messages[]` → proxy → Foundry → assistant text → render. **No function key**, **no shared secret in the SPFx bundle** — Easy Auth + the SPFx-acquired Entra token are the only auth gate.

## Run locally

```bash
cd webparts/mcp365-foundry-chat
npm install
npm run deploy         # provisions the proxy + auto-wires serve.json
npm start              # workbench opens with the property pane pre-filled
```

Send a message. If you get a reply, the full auth chain works: AAD bearer → Easy Auth → Function App → managed identity → Foundry.

## Property pane

| Field | How it's filled |
|---|---|
| Backend URL | Auto-wired by `npm run deploy` |
| Backend API resource | Auto-wired by `npm run deploy` (e.g. `api://<guid>`) |

The model deployment name lives in the Function App's settings (env vars), not the property pane — swap models without touching the webpart.

## Other commands

```bash
npm run deploy:dry-run # walk through the prompts, see the plan, no Azure calls
npm run setup          # re-wires serve.json from .deploy-output.json (no Azure changes)
npm run teardown       # deletes the resource group; zero lingering infra
```

`npm run deploy:dry-run` is its own script on purpose — `npm run deploy --dry-run` doesn't work (npm consumes the flag).

## What it deliberately doesn't do

- **No MCP, no tools.** Those land in follow-up webparts in the same series.
- **No streaming.** Plain `await response.json()`. SSE through Functions on Consumption / Basic plans is unreliable; not the showcase point.
- **No history persistence.** Resets when the page reloads.
- **No system-prompt customisation in the UI.** It's a one-line default.
- **No function key in the browser.** Keys in a public client would be a shared secret. Easy Auth + the SPFx-acquired Entra token are the only auth gate.

The job is to prove the protected backend works. Everything else is a different post.
