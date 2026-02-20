# NExS Spreadsheet MCP App — V1 Design

> Aligned with `@modelcontextprotocol/ext-apps` v1.0.1 and `@modelcontextprotocol/sdk` (latest), spec version `2026-01-26`.

## Overview

An MCP App that renders live, interactive NExS spreadsheets inside Claude Web (and any other MCP Apps–capable host). The user gives Claude a published NExS URL, and the spreadsheet appears inline in the conversation — fully interactive, with all calculations handled natively by the NExS platform.

No authentication. No discovery API. The NExS platform publishes spreadsheets to public URLs by default — V1 simply wraps those URLs for in-chat display.

---

## Architecture

```
┌──────────────────────────────────────┐
│         NExS MCP App View            │  Sandboxed iframe in chat
│   (HTML + JS, embeds NExS iframe)    │
└──────────────────┬───────────────────┘
                   │ postMessage (JSON-RPC)
┌──────────────────▼───────────────────┐
│         MCP Host (Claude Web)        │  Renders View, routes messages
└──────────────────┬───────────────────┘
                   │ MCP Protocol (Streamable HTTP / stdio)
┌──────────────────▼───────────────────┐
│         NExS MCP Server              │  One tool, one UI resource
└──────────────────────────────────────┘
                   │
                   │ (no server-side calls — URL passed straight through)
                   ▼
┌──────────────────────────────────────┐
│  NExS Public URL (platform.nexs.com) │  Loaded directly by the nested iframe
└──────────────────────────────────────┘
```

The MCP server has exactly two registrations: one **tool** and one **UI resource**. There are no API calls, no tokens, no polling. The server's only job is to relay the user-provided URL into the sandboxed View.

---

## Server Implementation

### `server.ts` — Tool + Resource Registration

The server factory function creates a fresh `McpServer` instance. It uses `getUiCapability` to check whether the connected client supports MCP Apps before registering the UI-linked tool — this ensures graceful degradation to a text-only response on clients that don't support the extension.

```typescript
import {
  registerAppResource,
  registerAppTool,
  getUiCapability,
  RESOURCE_MIME_TYPE,
} from "@modelcontextprotocol/ext-apps/server";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import fs from "node:fs/promises";
import path from "node:path";

const DIST_DIR = path.join(import.meta.dirname, "dist");
const RESOURCE_URI = "ui://nexs/spreadsheet.html";

export function createServer(): McpServer {
  const server = new McpServer({
    name: "NExS Spreadsheet Viewer",
    version: "1.0.0",
  });

  // Register the tool with UI metadata.
  // When the host calls this tool, it reads _meta.ui.resourceUri
  // to know which resource to fetch and render as an interactive View.
  registerAppTool(
    server,
    "render_nexs_spreadsheet",
    {
      title: "Render NExS Spreadsheet",
      description:
        "Renders a live, interactive NExS spreadsheet in the conversation. " +
        "Use when the user provides a published NExS platform URL.",
      inputSchema: {
        app_url: z
          .string()
          .url()
          .describe("A published NExS spreadsheet URL (https://platform.nexs.com/...)."),
      },
      _meta: {
        ui: { resourceUri: RESOURCE_URI },
      },
    },
    async ({ app_url }) => ({
      content: [{ type: "text", text: `Rendering spreadsheet: ${app_url}` }],
      structuredContent: { app_url },
    })
  );

  // Register the UI resource — returns the bundled single-file HTML.
  // CSP frameDomains is declared in contents[]._meta.ui so the host
  // allows the nested NExS iframe.
  registerAppResource(
    server,
    "nexs-spreadsheet",
    RESOURCE_URI,
    { mimeType: RESOURCE_MIME_TYPE },
    async () => {
      const html = await fs.readFile(
        path.join(DIST_DIR, "spreadsheet.html"),
        "utf-8"
      );
      return {
        contents: [
          {
            uri: RESOURCE_URI,
            mimeType: RESOURCE_MIME_TYPE,
            text: html,
            _meta: {
              ui: {
                prefersBorder: true,
                csp: {
                  frameDomains: ["https://platform.nexs.com"],
                },
              },
            },
          },
        ],
      };
    }
  );

  return server;
}
```

### `main.ts` — Entry Point (Dual Transport)

Supports both Streamable HTTP (for Claude Web via tunnel) and stdio (for Claude Desktop / VS Code). Uses `createMcpExpressApp` from the SDK.

```typescript
import { createMcpExpressApp } from "@modelcontextprotocol/sdk/server/express.js";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import cors from "cors";
import type { Request, Response } from "express";
import { createServer } from "./server.js";

async function startStreamableHTTPServer(
  factory: () => McpServer
): Promise<void> {
  const port = parseInt(process.env.PORT ?? "3001", 10);

  const app = createMcpExpressApp({ host: "0.0.0.0" });
  app.use(cors());

  app.all("/mcp", async (req: Request, res: Response) => {
    const server = factory();
    const transport = new StreamableHTTPServerTransport({
      sessionIdGenerator: undefined,
    });

    res.on("close", () => {
      transport.close().catch(() => {});
      server.close().catch(() => {});
    });

    try {
      await server.connect(transport);
      await transport.handleRequest(req, res, req.body);
    } catch (error) {
      console.error("MCP error:", error);
      if (!res.headersSent) {
        res.status(500).json({
          jsonrpc: "2.0",
          error: { code: -32603, message: "Internal server error" },
          id: null,
        });
      }
    }
  });

  const httpServer = app.listen(port, (err) => {
    if (err) {
      console.error("Failed to start server:", err);
      process.exit(1);
    }
    console.log(`MCP server listening on http://localhost:${port}/mcp`);
  });

  const shutdown = () => {
    console.log("\nShutting down...");
    httpServer.close(() => process.exit(0));
  };
  process.on("SIGINT", shutdown);
  process.on("SIGTERM", shutdown);
}

async function startStdioServer(factory: () => McpServer): Promise<void> {
  await factory().connect(new StdioServerTransport());
}

async function main() {
  if (process.argv.includes("--stdio")) {
    await startStdioServer(createServer);
  } else {
    await startStreamableHTTPServer(createServer);
  }
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
```

---

## View Implementation

### `spreadsheet.html` — UI Shell

```html
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <title>NExS Spreadsheet</title>
  <style>
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
    html, body { height: 100%; width: 100%; overflow: hidden; }
    .container {
      width: 100%;
      height: 100%;
      display: flex;
      align-items: center;
      justify-content: center;
      background: var(--color-background-primary, #f8f9fa);
      font-family: var(--font-sans, system-ui, -apple-system, sans-serif);
      color: var(--color-text-primary, #374151);
    }
    .loading { font-size: 14px; }
    iframe { width: 100%; height: 100%; border: 0; }
  </style>
</head>
<body>
  <div class="container" id="app-root">
    <p class="loading">Loading NExS spreadsheet…</p>
  </div>
  <script type="module" src="/src/spreadsheet.ts"></script>
</body>
</html>
```

### `src/spreadsheet.ts` — View Logic

Uses the `App` class from `@modelcontextprotocol/ext-apps`. Registers the `ontoolresult` handler before calling `connect()` so the initial result isn't missed. Applies host theme variables for visual consistency.

```typescript
import { App } from "@modelcontextprotocol/ext-apps";
import type { McpUiHostContext } from "@modelcontextprotocol/ext-apps";
import {
  applyDocumentTheme,
  applyHostStyleVariables,
  applyHostFonts,
} from "@modelcontextprotocol/ext-apps";

const root = document.getElementById("app-root")!;

const app = new App({ name: "NExS Spreadsheet Viewer", version: "1.0.0" });

// --- Theme integration ---
function applyHostContext(ctx: Partial<McpUiHostContext>) {
  if (ctx.theme) applyDocumentTheme(ctx.theme);
  if (ctx.styles?.variables) applyHostStyleVariables(ctx.styles.variables);
  if (ctx.styles?.css?.fonts) applyHostFonts(ctx.styles.css.fonts);
}

app.onhostcontextchanged = applyHostContext;

// --- Handle the tool result pushed by the host ---
app.ontoolresult = (result) => {
  const data = result.structuredContent as { app_url?: string };
  const url = data?.app_url;

  if (!url || !url.startsWith("https://platform.nexs.com/")) {
    root.innerHTML = `<p class="loading">Invalid or missing spreadsheet URL.</p>`;
    return;
  }

  root.innerHTML = `<iframe src="${url}" allowfullscreen></iframe>`;
};

// --- Connect to host (must be after handler registration) ---
app.connect().then(() => {
  const ctx = app.getHostContext();
  if (ctx) applyHostContext(ctx);
});
```

---

## Project Structure

```
nexs-mcp-app/
├── package.json
├── tsconfig.json              # Type-checking (noEmit, includes src + server)
├── tsconfig.server.json       # Server compilation (NodeNext, declaration only)
├── vite.config.ts             # Bundles View into single HTML via vite-plugin-singlefile
├── main.ts                    # Entry point — dual transport (HTTP + stdio)
├── server.ts                  # Tool + resource registration (factory function)
├── spreadsheet.html           # View shell
├── src/
│   └── spreadsheet.ts         # View logic (App class, theme, iframe mount)
└── dist/                      # Build output
```

## `package.json`

```json
{
  "type": "module",
  "scripts": {
    "build": "tsc --noEmit && tsc -p tsconfig.server.json && cross-env INPUT=spreadsheet.html vite build",
    "start": "concurrently \"cross-env NODE_ENV=development INPUT=spreadsheet.html vite build --watch\" \"tsx watch main.ts\"",
    "serve": "npm run build && node dist/main.js"
  },
  "dependencies": {
    "@modelcontextprotocol/ext-apps": "^1.0.1",
    "@modelcontextprotocol/sdk": "latest",
    "cors": "^2.8.5",
    "express": "^5.2.1",
    "zod": "^3.25.0"
  },
  "devDependencies": {
    "@types/cors": "^2.8.17",
    "@types/express": "^5.0.0",
    "@types/node": "^22.0.0",
    "concurrently": "^9.0.0",
    "cross-env": "^7.0.3",
    "tsx": "^4.0.0",
    "typescript": "^5.7.0",
    "vite": "^6.0.0",
    "vite-plugin-singlefile": "^2.0.0"
  }
}
```

## TypeScript Configuration

### `tsconfig.json` (type-checking, noEmit)

```json
{
  "compilerOptions": {
    "target": "ESNext",
    "lib": ["ESNext", "DOM", "DOM.Iterable"],
    "module": "ESNext",
    "moduleResolution": "bundler",
    "allowImportingTsExtensions": true,
    "resolveJsonModule": true,
    "isolatedModules": true,
    "verbatimModuleSyntax": true,
    "noEmit": true,
    "strict": true,
    "skipLibCheck": true,
    "noUnusedLocals": true,
    "noUnusedParameters": true,
    "noFallthroughCasesInSwitch": true
  },
  "include": ["src", "server.ts", "main.ts"]
}
```

### `tsconfig.server.json` (server compilation)

```json
{
  "compilerOptions": {
    "target": "ES2022",
    "lib": ["ES2022"],
    "module": "NodeNext",
    "moduleResolution": "NodeNext",
    "declaration": true,
    "emitDeclarationOnly": true,
    "outDir": "./dist",
    "rootDir": ".",
    "strict": true,
    "esModuleInterop": true,
    "skipLibCheck": true,
    "noUnusedLocals": true,
    "noUnusedParameters": true,
    "noFallthroughCasesInSwitch": true
  },
  "include": ["server.ts", "main.ts"]
}
```

### `vite.config.ts`

```typescript
import { defineConfig } from "vite";
import { viteSingleFile } from "vite-plugin-singlefile";

const INPUT = process.env.INPUT;
if (!INPUT) throw new Error("INPUT environment variable is not set");

const isDevelopment = process.env.NODE_ENV === "development";

export default defineConfig({
  plugins: [viteSingleFile()],
  build: {
    sourcemap: isDevelopment ? "inline" : undefined,
    cssMinify: !isDevelopment,
    minify: !isDevelopment,
    rollupOptions: { input: INPUT },
    outDir: "dist",
    emptyOutDir: false,
  },
});
```

---

## Local Dev & Testing

### Claude Web (Streamable HTTP)

```bash
npm start                          # builds + starts on http://localhost:3001/mcp

# In a separate terminal, create a tunnel:
npx cloudflared tunnel --url http://localhost:3001

# In Claude: Settings → Connectors → Add Custom Connector
# Paste: https://<tunnel-id>.trycloudflare.com/mcp
```

### Claude Desktop (stdio)

Add to Claude Desktop's MCP server config:

```json
{
  "mcpServers": {
    "nexs-spreadsheet": {
      "command": "node",
      "args": ["<path-to>/nexs-mcp-app/dist/main.js", "--stdio"]
    }
  }
}
```

### ext-apps basic-host (standalone testing)

```bash
git clone https://github.com/modelcontextprotocol/ext-apps.git
cd ext-apps && npm install
cd examples/basic-host && npm start
# Open http://localhost:8080, select your server, call the tool
```

---

## Prerequisite

The NExS platform (`platform.nexs.com`) must allow iframe embedding from the host's sandbox origin. If NExS currently sends `X-Frame-Options: DENY` or a restrictive `frame-ancestors` CSP, that header needs to be relaxed for this to work.

---

## Future Work

- **Graceful degradation** — use `getUiCapability` to conditionally register the UI tool vs. a text-only fallback for non-MCP-Apps clients.
- **Fullscreen mode** — call `app.requestDisplayMode({ mode: "fullscreen" })` for a better spreadsheet editing experience.
- **Context feedback** — use `app.updateModelContext()` to push calculation results or selected cell data back into Claude's conversation context.
- **Spreadsheet discovery** — add a resource listing available apps so Claude can resolve names to URLs.
- **Per-user auth** — OAuth so different users see their own NExS workspaces.
- **Directory submission** — `frameDomains` triggers higher review scrutiny; may need to render natively instead of embedding an iframe.
