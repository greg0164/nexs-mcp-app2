import {
  registerAppResource,
  registerAppTool,
  RESOURCE_MIME_TYPE,
} from "@modelcontextprotocol/ext-apps/server";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { CallToolResult, ReadResourceResult } from "@modelcontextprotocol/sdk/types.js";
import fs from "node:fs/promises";
import path from "node:path";
import { z } from "zod";

// Works both from source (server.ts via tsx) and compiled (dist/server.js)
const DIST_DIR = import.meta.filename.endsWith(".ts")
  ? path.join(import.meta.dirname, "dist")
  : import.meta.dirname;

const RESOURCE_URI = "ui://nexs/spreadsheet.html";

// Module-level store — survives across the per-request McpServer instances
// because this module is loaded once per process. Used by the restore tool
// so the App View can recover the URL after a page refresh.
let lastSpreadsheetUrl: string | null = null;

/**
 * Creates a new MCP server instance with the NExS spreadsheet tool and
 * its corresponding UI resource registered.
 */
export function createServer(): McpServer {
  const server = new McpServer({
    name: "NExS Spreadsheet Viewer",
    version: "1.0.0",
  });

  // Register the tool with UI metadata.
  // When the host calls this tool, it reads `_meta.ui.resourceUri` to know
  // which resource to fetch and render as an interactive View.
  registerAppTool(
    server,
    "render_nexs_spreadsheet",
    {
      title: "Render NExS Spreadsheet",
      description:
        "Renders a live, interactive NExS spreadsheet inline in the conversation. " +
        "Use when the user provides a published NExS platform URL " +
        "(https://platform.nexs.com/...).",
      inputSchema: {
        app_url: z
          .string()
          .url()
          .describe(
            "A published NExS spreadsheet URL (https://platform.nexs.com/...)."
          ),
      },
      // outputSchema causes the host to include structuredContent in the
      // ui/notifications/tool-result sent to the App View.
      outputSchema: {
        app_url: z
          .string()
          .url()
          .describe("The NExS spreadsheet URL being rendered."),
      },
      _meta: {
        ui: { resourceUri: RESOURCE_URI },
      },
    },
    async ({ app_url }): Promise<CallToolResult> => {
      lastSpreadsheetUrl = app_url;
      return {
        content: [{ type: "text", text: `Rendering NExS spreadsheet: ${app_url}` }],
        structuredContent: { app_url },
      };
    }
  );

  // Internal restore tool — callable only by the App View (visibility: ["app"]),
  // invisible to the LLM. The App calls this on page refresh to recover the
  // last-rendered URL when ontoolresult is not re-delivered by the host.
  registerAppTool(
    server,
    "restore_nexs_spreadsheet",
    {
      description: "Returns the last-rendered NExS spreadsheet URL for refresh recovery.",
      outputSchema: {
        app_url: z.string().url().nullable().describe("Last spreadsheet URL, or null."),
      },
      _meta: {
        ui: {
          resourceUri: RESOURCE_URI,
          visibility: ["app"],
        },
      },
    },
    async (): Promise<CallToolResult> => ({
      content: [],
      structuredContent: { app_url: lastSpreadsheetUrl },
    })
  );

  // Register the UI resource — returns the single-file bundled HTML View.
  // The CSP `frameDomains` entry allows the View to embed the NExS iframe.
  registerAppResource(
    server,
    "NExS Spreadsheet View",
    RESOURCE_URI,
    { mimeType: RESOURCE_MIME_TYPE },
    async (): Promise<ReadResourceResult> => {
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
