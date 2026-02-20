# NExS Spreadsheet MCP App

An MCP App extension that renders live, interactive **NExS spreadsheets** inline inside Claude Web (and any MCP Apps–capable host). The user gives Claude a published NExS URL, and the spreadsheet appears directly in the conversation — fully interactive, with all calculations handled by the NExS platform.

Built with [`@modelcontextprotocol/ext-apps`](https://github.com/modelcontextprotocol/ext-apps) v1.0.1.

---

## How It Works

```
User gives Claude a platform.nexs.com URL
        ↓
Claude calls render_nexs_spreadsheet tool
        ↓
MCP server returns tool result + HTML resource URI
        ↓
Host (Claude Web) fetches and renders the View
        ↓
View mounts the NExS URL in a sandboxed iframe
```

The server itself does no API calls — it just relays the user-provided URL into the sandboxed View. The NExS iframe loads directly from `platform.nexs.com`.

---

## Project Structure

```
nexs-mcp-app/
├── package.json
├── tsconfig.json              # Type-checking (noEmit), includes src + server + main
├── tsconfig.server.json       # Server compilation (NodeNext) → dist/
├── vite.config.ts             # Bundles View into single-file HTML via vite-plugin-singlefile
├── main.ts                    # Entry point — dual transport (HTTP + stdio)
├── server.ts                  # Tool + resource registration (factory function)
├── spreadsheet.html           # View shell (Vite entry point)
├── src/
│   └── spreadsheet.ts         # View logic (App class, theme, iframe mount)
└── dist/                      # Build output (gitignored)
```

---

## Getting Started

### Prerequisites

- Node.js 18+
- npm 9+

### Install

```bash
npm install
```

### Development (hot-reload)

```bash
npm start
# Starts vite --watch for the View and tsx watch for the server
# MCP endpoint: http://localhost:3001/mcp
```

### Production build + serve

```bash
npm run serve
# Compiles TypeScript, bundles the View, then starts the HTTP server
# MCP endpoint: http://localhost:3001/mcp
```

---

## Connecting to Claude

### Claude Web (Streamable HTTP)

1. Run the server: `npm run serve`
2. Create a tunnel: `npx cloudflared tunnel --url http://localhost:3001`
3. In Claude: **Settings → Connectors → Add Custom Connector**
4. Paste: `https://<tunnel-id>.trycloudflare.com/mcp`

### Claude Desktop / VS Code (stdio)

Add to your MCP server config:

```json
{
  "mcpServers": {
    "nexs-spreadsheet": {
      "command": "node",
      "args": ["<path-to-repo>/dist/main.js", "--stdio"]
    }
  }
}
```

### Local testing with ext-apps basic-host

```bash
git clone https://github.com/modelcontextprotocol/ext-apps.git /tmp/mcp-ext-apps
cd /tmp/mcp-ext-apps/examples/basic-host && npm install && npm start
# Open http://localhost:8080
```

---

## Usage

Once connected, ask Claude:

> "Show me this NExS spreadsheet: https://platform.nexs.com/your-sheet-id"

Claude will call `render_nexs_spreadsheet` and the spreadsheet appears inline in the conversation.

---

## Environment Variables

| Variable | Default | Description |
|---|---|---|
| `PORT` | `3001` | HTTP port for Streamable HTTP transport |

---

## Future Work

- **Graceful degradation** — use `getUiCapability` to register a text-only fallback tool for non-MCP-Apps clients
- **Fullscreen mode** — `app.requestDisplayMode({ mode: "fullscreen" })` for better spreadsheet editing
- **Context feedback** — `app.updateModelContext()` to push cell data/results into Claude's context
- **Spreadsheet discovery** — resource listing available NExS apps by name
- **Per-user auth** — OAuth so different users see their own NExS workspaces

---

## Prerequisite

The NExS platform (`platform.nexs.com`) must allow iframe embedding from the host's sandbox origin. If NExS sends `X-Frame-Options: DENY` or a restrictive `frame-ancestors` CSP, that header needs to be relaxed for the embedded iframe to load.
