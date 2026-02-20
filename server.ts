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
      _meta: {
        ui: { resourceUri: RESOURCE_URI },
      },
    },
    async ({ app_url }): Promise<CallToolResult> => ({
      content: [{ type: "text", text: `Rendering NExS spreadsheet: ${app_url}` }],
      structuredContent: { app_url },
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
