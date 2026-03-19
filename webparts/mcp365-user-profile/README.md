# MCP365 Explorer: User Profile

Interactive SPFx webpart for exploring the **mcp_MeServer** — the M365 MCP User Profile server with 5 tools for user profiles, manager chains, and org hierarchy.

![MCP365 Explorer: User Profile](assets/mcp_explorer_profile.gif)

## What it does

Connect directly to the Agent 365 MCP gateway from the browser — no backend required — and interactively explore the User Profile server:

- **Server info**: protocol version, capabilities, session ID
- **Tool browser**: all 5 tools with live schemas from `tools/list`
- **Auto-generated forms**: parameter fields built at runtime from the server's `inputSchema`
- **Formatted responses**: embedded JSON extracted and prettified, Graph noise stripped
- **Searchable log viewer**: every JSON-RPC exchange with sorting and expand
- **Custom presets**: save your own parameter sets to browser localStorage

## Tools

| Tool | Description |
|------|-------------|
| GetMyDetails | Current user's profile |
| GetUserDetails | Lookup by UPN or Entra ID |
| GetMultipleUsersDetails | Search users by name, title, location |
| GetManagerDetails | User's manager |
| GetDirectReportsDetails | User's direct reports |

## Prerequisites

1. **Agents Toolkit Preview** — tenant enrolled in the Microsoft 365 Agents Toolkit program
2. **Service Principal** — run `scripts/New-Agent365ServicePrincipal.ps1` (one-time admin operation)
3. **Environment ID** — Power Platform environment GUID from [admin center](https://admin.powerplatform.microsoft.com/)
4. **Node.js 22+** and SPFx 1.22 development environment

## Build & Deploy

```bash
cd webparts/mcp365-user-profile
npm install
npx heft build --clean

# Package for production
npx heft test --clean --production
npx heft package-solution --production
```

Upload `sharepoint/solution/mcp365-user-profile.sppkg` to your app catalog, then approve the **McpServers.Me.All** permission in SharePoint admin center > API Management.

Add the webpart to a page and configure the **Environment ID** in the property pane.

## Part of MCP365 Explorer

This is the first webpart in the [MCP365 Explorer](https://github.com/ferrarirosso/mcp365-explorer) series — one webpart per M365 MCP server, each with a matching [blog post](https://www.puntobello.ch/en/nello/mcp365_explorer_intro/).
