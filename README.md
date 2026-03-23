# MCP365 Explorer

## Compatibility

![SPFx 1.22](https://img.shields.io/badge/SPFx-1.22.1-green.svg)
![Node.js v22](https://img.shields.io/badge/Node.js-v22-green.svg)
![Compatible with SharePoint Online](https://img.shields.io/badge/SharePoint%20Online-Compatible-green.svg)
![Does not work with SharePoint 2019](https://img.shields.io/badge/SharePoint%20Server%202019-Incompatible-red.svg "SharePoint Server 2019 requires SPFx 1.4.1 or lower")
![Does not work with SharePoint 2016 (Feature Pack 2)](https://img.shields.io/badge/SharePoint%20Server%202016%20(Feature%20Pack%202)-Incompatible-red.svg "SharePoint Server 2016 Feature Pack 2 requires SPFx 1.1")
![Teams Compatible](https://img.shields.io/badge/Teams-Compatible-green.svg)
![Hosted Workbench Compatible](https://img.shields.io/badge/Hosted%20Workbench-Compatible-green.svg)

## Summary

Open-source SPFx webparts for exploring and testing the Microsoft 365 [Work IQ](https://learn.microsoft.com/en-us/microsoft-agent-365/tooling-servers-overview) MCP servers — Microsoft's umbrella for all MCP capabilities in the M365 ecosystem. One webpart per server, each accompanied by a [blog post](https://www.puntobello.ch/en/nello/mcp365_explorer_intro/).

**Key finding:** You can call the Work IQ MCP servers **directly from an SPFx webpart** — no backend, no Azure Functions proxy, no additional infrastructure. Just `fetch`, a bearer token from `AadTokenProvider`, and the JSON-RPC protocol.

> **Preview notice:** The Agents 365 Tools and MCP servers are part of the [Microsoft Agent 365 tooling servers preview](https://learn.microsoft.com/en-us/microsoft-agent-365/tooling-servers-overview). These features are in preview, may change, and should not be used in production workloads.

## Webparts

| Server | ID | Tools | Status |
|--------|-----|-------|--------|
| [Work IQ User](webparts/mcp365-user-profile/) | `mcp_MeServer` | 5 | Available |
| [Work IQ SharePoint](webparts/mcp365-sharepoint-lists/) | `mcp_SharePointRemoteServer` | 35 | Available |
| Work IQ Calendar | `mcp_CalendarTools` | 13 | Available soon |
| Work IQ Mail | `mcp_MailTools` | 21 | Planned |
| Work IQ Teams | `mcp_TeamsServer` | 26 | Planned |
| Work IQ OneDrive | `mcp_OneDriveRemoteServer` | — | Planned |
| Work IQ Word | `mcp_WordServer` | 4 | Planned |

## What Each Webpart Does

- **Showcase mode** — click a button, see the result. No JSON, no parameters.
- **Explorer mode** — browse tools, inspect live schemas from `tools/list`, auto-generated parameter forms, formatted responses, searchable log viewer
- **Custom presets** — save your own parameter sets to browser localStorage

![MCP365 Explorer: User Profile](webparts/mcp365-user-profile/assets/mcp_explorer_profile.gif)

## Prerequisites

1. **Microsoft Frontier AI Program** — [Enrollment](https://adoption.microsoft.com/en-us/copilot/frontier-program/)
2. **Work IQ Tools Service Principal** — Run [`New-Agent365ToolsServicePrincipalProdPublic.ps1`](scripts/New-Agent365ServicePrincipal.ps1) (one-time admin operation). See [Microsoft's guide](https://learn.microsoft.com/en-us/microsoft-agent-365/developer/tooling#set-up-service-principal).
3. **Power Platform Environment ID** — From [Power Platform admin center](https://admin.powerplatform.microsoft.com/). See [how to find it](https://learn.microsoft.com/en-us/power-platform/admin/determine-org-id-name).
4. **Node.js 22+** and SPFx 1.22

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

## Blog Series

- [MCP365 Explorer — Introduction + User Profile](https://www.puntobello.ch/en/nello/mcp365_explorer_intro/)
- [MCP365 Explorer — Work IQ SharePoint: 35 tools](https://www.puntobello.ch/en/nello/mcp365_explorer_sharepoint_lists/)
- More posts coming — one per server

## Resources

- [Work IQ MCP Servers Overview](https://learn.microsoft.com/en-us/microsoft-agent-365/tooling-servers-overview)
- [Model Context Protocol Specification](https://modelcontextprotocol.io/)
- [GriMoire — Visual AI Assistant for M365](https://grimoire-hie.github.io/)
- [SPFx Documentation](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/sharepoint-framework-overview)
