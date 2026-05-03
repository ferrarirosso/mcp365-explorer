# MCP365 Explorer

## Compatibility

![SPFx 1.22](https://img.shields.io/badge/SPFx-1.22.1-green.svg)
![Node.js v22](https://img.shields.io/badge/Node.js-v22-green.svg)
![Compatible with SharePoint Online](https://img.shields.io/badge/SharePoint%20Online-Compatible-green.svg)
![Does not work with SharePoint 2019](https://img.shields.io/badge/SharePoint%20Server%202019-Incompatible-red.svg "SharePoint Server 2019 requires SPFx 1.4.1 or lower")
![Does not work with SharePoint 2016 (Feature Pack 2)](https://img.shields.io/badge/SharePoint%20Server%202016%20(Feature%20Pack%202)-Incompatible-red.svg "SharePoint Server 2016 Feature Pack 2 requires SPFx 1.1")
![Teams Compatible](https://img.shields.io/badge/Teams-Compatible-green.svg)
![Hosted Workbench Compatible](https://img.shields.io/badge/Hosted%20Workbench-Compatible-green.svg)

[![Open in GitHub Codespaces](https://github.com/codespaces/badge.svg)](https://codespaces.new/ferrarirosso/mcp365-explorer)

## Summary

Open-source SPFx webparts for exploring Microsoft 365's MCP and Foundry surface. Two patterns:

- **[Work IQ](https://learn.microsoft.com/en-us/microsoft-agent-365/tooling-servers-overview) MCP servers, called directly from an SPFx webpart** — no backend, no proxy, just `fetch` + `AadTokenProvider` + JSON-RPC. One webpart per server.
- **Microsoft Foundry behind a protected Function App** — keyless backend, Easy Auth + managed identity, deployed in one command via [`spfx-foundry-deploy`](https://github.com/ferrarirosso/spfx-foundry-deploy). Two webparts: `mcp365-foundry-chat` (chat-only auth-chain showcase) and `mcp365-lists-chat` (chat + SharePoint Lists MCP — the first **agentic** webpart in the series).

Each webpart pairs with a [blog post](https://www.puntobello.ch/en/nello/mcp365_explorer_intro/).

> **Preview notice:** The Agents 365 Tools and MCP servers are part of the [Microsoft Agent 365 tooling servers preview](https://learn.microsoft.com/en-us/microsoft-agent-365/tooling-servers-overview). These features are in preview, may change, and should not be used in production workloads.

## Webparts

### Direct from the browser

No backend. AAD bearer + JSON-RPC straight to the Work IQ MCP server.

| Server | ID | Tools | Status |
|--------|-----|-------|--------|
| [Work IQ User](webparts/mcp365-user-profile/) | `mcp_MeServer` | 5 | Available |
| [Work IQ SharePoint](webparts/mcp365-sharepoint-lists/) | `mcp_SharePointRemoteServer` | 35 | Available |
| [Work IQ Calendar](webparts/mcp365-calendar/) | `mcp_CalendarTools` | 13 | Available |
| [Work IQ Mail](webparts/mcp365-mail/) | `mcp_MailTools` | 22 | Available |
| [Work IQ Teams](webparts/mcp365-teams/) | `mcp_TeamsServer` | 28 | Available |
| [Work IQ OneDrive](webparts/mcp365-onedrive/) | `mcp_OneDriveRemoteServer` | 13 | Available |
| [Work IQ Word](webparts/mcp365-word/) | `mcp_WordServer` | 4 | Available |

### Through a protected Function App proxy

Foundry-backed Azure Function App, keyless to Foundry, Easy Auth-protected from the browser. Provisioned with [`spfx-foundry-deploy`](https://github.com/ferrarirosso/spfx-foundry-deploy).

| Webpart | Backend | Purpose | Status |
|---------|---------|---------|--------|
| [mcp365-foundry-chat](webparts/mcp365-foundry-chat/) | Foundry (chat-completions) | Chat-only showcase — proves the deployment + auth chain end-to-end | Available |
| [mcp365-lists-chat](webparts/mcp365-lists-chat/) | Foundry + Work IQ SharePoint MCP | Agentic chat — LLM picks tools from `tools/list` and executes them via MCP | Available |

## What Each Webpart Does

- **Showcase mode** — click a button, see the result. No JSON, no parameters.
- **Explorer mode** — browse tools, inspect live schemas from `tools/list`, auto-generated parameter forms, formatted responses, searchable log viewer
- **Custom presets** — save your own parameter sets to browser localStorage

![MCP365 Explorer: User Profile](webparts/mcp365-user-profile/assets/mcp_explorer_profile.gif)

## Prerequisites

**For the Work IQ webparts:**

1. **Microsoft Frontier AI Program** — [Enrollment](https://adoption.microsoft.com/en-us/copilot/frontier-program/)
2. **Work IQ Tools Service Principal** — Run [`New-Agent365ServicePrincipal.ps1`](scripts/New-Agent365ServicePrincipal.ps1) (one-time admin operation). See [Microsoft's guide](https://learn.microsoft.com/en-us/microsoft-agent-365/developer/tooling#set-up-service-principal).
3. **Power Platform Environment ID** — From [Power Platform admin center](https://admin.powerplatform.microsoft.com/). See [how to find it](https://learn.microsoft.com/en-us/power-platform/admin/determine-org-id-name).
4. **Node.js 22+** and SPFx 1.22

**For the Foundry-backed webparts — additionally:**

5. **Azure subscription** with quota for `gpt-5-mini` (or your chosen model) in the target region
6. **Azure CLI ≥ 2.69** and **Azure Functions Core Tools v4** (`func`) — both pre-installed in the [included Codespaces devcontainer](.devcontainer/devcontainer.json)

> **Note on service principal naming:** These webparts use `"resource": "Work IQ Tools"` in `package-solution.json`. If you ran the service principal script **before March 12, 2026**, your enterprise app is named "Agent 365 Tools" instead. In that case, change `"resource": "Work IQ Tools"` to `"resource": "Agent 365 Tools"` in each webpart's `config/package-solution.json` — otherwise the SharePoint admin center will reject the API permission request.

## Quick Start

### User Profile (5 tools)

```bash
cd webparts/mcp365-user-profile
npm install
npx heft build --clean
npx heft test --clean --production && npx heft package-solution --production
```

Upload `.sppkg`, approve `McpServers.Me.All`, add to a page.

### SharePoint (35 tools)

```bash
cd webparts/mcp365-sharepoint-lists
npm install
npx heft build --clean
npx heft test --clean --production && npx heft package-solution --production
```

Upload `.sppkg`, approve `McpServers.SharePoint.All`, add to a page.

### Foundry chat showcase

```bash
cd webparts/mcp365-foundry-chat
npm install
npm run deploy        # provisions the proxy + auto-wires serve.json (~5 min)
npm start             # workbench opens with the property pane pre-filled
```

`npm run deploy` walks you through a one-screen review form (region, model, prefix), then provisions everything: AI Foundry resource + model deployment, Storage Account, Function App with managed identity, Backend API Entra app, Easy Auth, platform hardening, App Insights. No function key in the browser. Full breakdown at [`spfx-foundry-deploy`](https://github.com/ferrarirosso/spfx-foundry-deploy).

`npm run teardown` removes everything — resource group, soft-delete purge, Entra app — so you can experiment without lingering infra.

### Lists chat (agentic — Foundry + Work IQ SharePoint MCP)

```bash
cd webparts/mcp365-lists-chat
npm install
npm run deploy        # provisions the proxy + auto-wires serve.json (~5 min)
npm start             # workbench opens with backendUrl + backendApiResource pre-filled
```

After deploy, set the **Environment ID** in the property pane (Power Platform env GUID — `pac admin list`). Approve `McpServers.SharePoint.All` in SharePoint admin centre to grant the MCP scope. Same deployer as `mcp365-foundry-chat`; the difference is what the chat is connected to.

## Blog Series

- [MCP365 Explorer — Introduction + User Profile](https://www.puntobello.ch/en/nello/mcp365_explorer_intro/)
- [MCP365 Explorer — Work IQ SharePoint: 35 tools](https://www.puntobello.ch/en/nello/mcp365_explorer_sharepoint_lists/)
- [MCP365 Explorer — Work IQ Calendar: events, meetings, and availability](https://www.puntobello.ch/en/nello/mcp365_explorer_calendar/)
- [MCP365 Explorer — Work IQ Mail: search, draft, send, and reply](https://www.puntobello.ch/en/nello/mcp365_explorer_mail/)
- [MCP365 Explorer — Work IQ Teams: teams, channels, and messages](https://www.puntobello.ch/en/nello/mcp365_explorer_teams/)
- [MCP365 Explorer — From buttons to language: chat with the SharePoint Lists MCP server](https://www.puntobello.ch/en/nello/mcp365_explorer_lists_chat/)
- More posts coming — agentic workflows across the 7 Work IQ servers

## Resources

- [Work IQ MCP Servers Overview](https://learn.microsoft.com/en-us/microsoft-agent-365/tooling-servers-overview)
- [Model Context Protocol Specification](https://modelcontextprotocol.io/)
- [`spfx-foundry-deploy`](https://github.com/ferrarirosso/spfx-foundry-deploy) — the Foundry-backed deployer
- [GriMoire](https://github.com/grimoire-hie)
- [SPFx Documentation](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/sharepoint-framework-overview)
