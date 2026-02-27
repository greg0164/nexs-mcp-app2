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

// ---------------------------------------------------------------------------
// NExS interact API types
// ---------------------------------------------------------------------------

interface NexsCellInfo {
  addr: string;
  data: number | string;
  datatype: "numeric" | "string" | "error" | "n/a";
  text: string;
  formula?: string;
}

interface NexsView {
  name: string;
  sheetName: string;
  range: string;
  isInvisible: boolean;
}

/** [sheetname, cellinfo] tuple as returned by both init and interact. */
type NexsCellEntry = [string, NexsCellInfo];

/**
 * Session obtained from the NExS init API when the spreadsheet was loaded.
 * get_cell syncs deltas via interact before every read; set_cell writes via
 * interact and applies returned deltas.  The session UUID never changes within
 * a conversation — it is the same session the iframe is displaying.
 */
interface NexsSession {
  appUuid: string;
  sessionId: string;
  revision: number;
  views: NexsView[];
  /** Full current-state cache keyed by "SheetName!ADDR" (addr uppercased). */
  cellCache: Map<string, { sheetName: string; ci: NexsCellInfo }>;
}

// ---------------------------------------------------------------------------
// Module-level store — one process, survives across per-request McpServer
// instances.
// ---------------------------------------------------------------------------

let lastSpreadsheetUrl: string | null = null;
let nexsSession: NexsSession | null = null;

// ---------------------------------------------------------------------------
// NExS API helpers
// ---------------------------------------------------------------------------

function extractNexsUuid(url: string): string | null {
  const m = url.match(
    /\/([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})/i
  );
  return m ? m[1] : null;
}

function parseCellRef(cellRef: string): { sheet: string | null; cell: string } {
  const bangIdx = cellRef.indexOf("!");
  if (bangIdx !== -1) {
    return { sheet: cellRef.slice(0, bangIdx), cell: cellRef.slice(bangIdx + 1) };
  }
  return { sheet: null, cell: cellRef };
}

interface NexsInitResult {
  sessionId: string;
  revision: number;
  views: NexsView[];
  values: NexsCellEntry[];
}

/**
 * POST /api/app/{uuid}/init
 *
 * Creates a NExS session and returns all current cell values.
 * The endpoint is csrf_exempt and works without cookies for public apps.
 * The returned session UUID is the one used by the iframe for the lifetime of
 * the conversation; get_cell / set_cell both operate on this same session.
 */
async function nexsInit(appUuid: string): Promise<NexsInitResult> {
  const resp = await fetch(
    `https://platform.nexs.com/api/app/${appUuid}/init`,
    {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({}),
    }
  );
  if (!resp.ok) {
    const text = await resp.text().catch(() => "");
    throw new Error(`NExS init failed (${resp.status}): ${text}`);
  }
  const data = (await resp.json()) as {
    session: string;
    revision: number;
    views: NexsView[];
    values: NexsCellEntry[];
  };
  return {
    sessionId: data.session,
    revision: data.revision,
    views: data.views,
    values: data.values,
  };
}

interface NexsInteractResult {
  revision: number;
  values: NexsCellEntry[];
}

/**
 * POST /api/app/{uuid}/interact
 *
 * Operates on an existing session.  Returns only cells that changed since the
 * given revision (delta).  Pass inputs=[] for a read-only sync; pass
 * [[sheetname, celladdr, value], ...] to write cells.
 */
async function nexsInteract(
  appUuid: string,
  sessionId: string,
  revision: number,
  inputs: [string, string, string | number][],
): Promise<NexsInteractResult> {
  const resp = await fetch(
    `https://platform.nexs.com/api/app/${appUuid}/interact`,
    {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ session: sessionId, revision, inputs }),
    }
  );
  if (!resp.ok) {
    const text = await resp.text().catch(() => "");
    throw new Error(`NExS interact failed (${resp.status}): ${text}`);
  }
  return resp.json() as Promise<NexsInteractResult>;
}

function applyDelta(session: NexsSession, result: NexsInteractResult): void {
  for (const [sn, ci] of result.values) {
    session.cellCache.set(`${sn}!${ci.addr.toUpperCase()}`, { sheetName: sn, ci });
  }
  session.revision = result.revision;
}

function buildCellCache(values: NexsCellEntry[]): NexsSession["cellCache"] {
  const cache: NexsSession["cellCache"] = new Map();
  for (const [sn, ci] of values) {
    cache.set(`${sn}!${ci.addr.toUpperCase()}`, { sheetName: sn, ci });
  }
  return cache;
}

// ---------------------------------------------------------------------------
// MCP server
// ---------------------------------------------------------------------------

export function createServer(): McpServer {
  const server = new McpServer({
    name: "NExS Spreadsheet Viewer",
    version: "1.0.0",
  });

  // ---------------------------------------------------------------------------
  // render_nexs_spreadsheet
  // ---------------------------------------------------------------------------
  registerAppTool(
    server,
    "render_nexs_spreadsheet",
    {
      title: "Render NExS Spreadsheet",
      description:
        "Displays a NExS spreadsheet as a live, interactive view in the conversation. " +
        "Call this ONCE when the user first provides a NExS URL. " +
        "NEVER call this again to read cell values — call get_cell for that. " +
        "Re-calling this reloads the iframe and resets all user edits.",
      inputSchema: {
        app_url: z
          .string()
          .url()
          .describe("A published NExS spreadsheet URL (https://platform.nexs.com/...)."),
      },
      outputSchema: {
        app_url: z.string().url().describe("The NExS spreadsheet URL being rendered."),
      },
      _meta: { ui: { resourceUri: RESOURCE_URI } },
    },
    async ({ app_url }): Promise<CallToolResult> => {
      lastSpreadsheetUrl = app_url;

      // Initialise the NExS session synchronously so get_cell / set_cell are
      // ready immediately after this tool returns.  The session UUID returned
      // here is the one the iframe will use for the lifetime of the conversation.
      const appUuid = extractNexsUuid(app_url);
      if (appUuid) {
        try {
          const init = await nexsInit(appUuid);
          nexsSession = {
            appUuid,
            sessionId: init.sessionId,
            revision: init.revision,
            views: init.views,
            cellCache: buildCellCache(init.values),
          };
        } catch (err) {
          // Non-fatal: render still succeeds; get_cell will surface the error.
          console.error("[NExS] Session init failed:", err);
        }
      }

      return {
        content: [
          {
            type: "text",
            text:
              `Rendering NExS spreadsheet: ${app_url}\n` +
              `Session ready. Use get_cell to read values and set_cell to write them.`,
          },
        ],
        structuredContent: { app_url },
      };
    }
  );

  // ---------------------------------------------------------------------------
  // restore_nexs_spreadsheet — internal, App View only.
  // ---------------------------------------------------------------------------
  registerAppTool(
    server,
    "restore_nexs_spreadsheet",
    {
      description: "Returns the last-rendered NExS spreadsheet URL for refresh recovery.",
      outputSchema: {
        app_url: z.string().url().nullable().describe("Last spreadsheet URL, or null."),
      },
      _meta: { ui: { resourceUri: RESOURCE_URI, visibility: ["app"] } },
    },
    async (): Promise<CallToolResult> => ({
      content: [],
      structuredContent: { app_url: lastSpreadsheetUrl },
    })
  );

  // ---------------------------------------------------------------------------
  // get_cell
  // ---------------------------------------------------------------------------
  server.registerTool(
    "get_cell",
    {
      title: "Get NExS Cell Value",
      description:
        "Reads the current value of a cell from the displayed NExS spreadsheet. " +
        "Use this — NOT render_nexs_spreadsheet — when the user asks about " +
        "quantities, values, or calculations shown in the spreadsheet. " +
        "Returns the formatted text, raw data value, and data type.",
      inputSchema: {
        cell_ref: z
          .string()
          .describe(
            "Cell address such as 'A1', 'B17', or 'Sheet1!A1'. " +
            "Named cells defined in the spreadsheet are also supported."
          ),
        sheet: z
          .string()
          .optional()
          .describe(
            "Sheet name. Optional when cell_ref includes the sheet (e.g. 'Sheet1!A1') " +
            "or the spreadsheet has only one sheet."
          ),
      },
      outputSchema: {
        sheet: z.string().describe("Sheet name the cell was found in."),
        addr: z.string().describe("Cell address (e.g. 'A1')."),
        value: z.union([z.string(), z.number()]).describe("Raw cell value."),
        text: z.string().describe("Formatted display text."),
        datatype: z.enum(["numeric", "string", "error", "n/a"]).describe("Cell data type."),
      },
    },
    async ({ cell_ref, sheet }): Promise<CallToolResult> => {
      if (!nexsSession) {
        return {
          isError: true,
          content: [
            { type: "text", text: "No NExS spreadsheet is loaded. Call render_nexs_spreadsheet first." },
          ],
        };
      }

      const { sheet: parsedSheet, cell: cellAddr } = parseCellRef(cell_ref);
      const sheetName =
        sheet ??
        parsedSheet ??
        nexsSession.views.find((v) => !v.isInvisible)?.sheetName ??
        null;

      // Sync: fetch any cells that changed since our last revision so the cache
      // reflects the current state of the spreadsheet session.
      try {
        const delta = await nexsInteract(
          nexsSession.appUuid,
          nexsSession.sessionId,
          nexsSession.revision,
          [],
        );
        applyDelta(nexsSession, delta);
      } catch (err) {
        return {
          isError: true,
          content: [
            {
              type: "text",
              text: `Failed to read from NExS: ${err instanceof Error ? err.message : String(err)}`,
            },
          ],
        };
      }

      // Search the cache for the requested cell.
      const cacheKey = sheetName
        ? `${sheetName}!${cellAddr.toUpperCase()}`
        : null;

      let found: { sheetName: string; ci: NexsCellInfo } | undefined;

      if (cacheKey) {
        found = nexsSession.cellCache.get(cacheKey);
      } else {
        const upper = cellAddr.toUpperCase();
        for (const [key, val] of nexsSession.cellCache) {
          if (key.endsWith(`!${upper}`)) {
            found = val;
            break;
          }
        }
      }

      if (!found) {
        const sheets = [...new Set([...nexsSession.cellCache.values()].map((v) => v.sheetName))];
        return {
          isError: true,
          content: [
            {
              type: "text",
              text:
                `Cell '${cell_ref}' not found. ` +
                `Known sheets: ${sheets.join(", ")}. ` +
                `Try specifying the sheet explicitly, e.g. '${sheets[0] ?? "Sheet1"}!${cellAddr}'.`,
            },
          ],
        };
      }

      const { sheetName: foundSheet, ci } = found;
      return {
        content: [
          { type: "text", text: `${foundSheet}!${ci.addr} = ${ci.text} (${ci.datatype})` },
        ],
        structuredContent: {
          sheet: foundSheet,
          addr: ci.addr,
          value: ci.data,
          text: ci.text,
          datatype: ci.datatype,
        },
      };
    }
  );

  // ---------------------------------------------------------------------------
  // set_cell
  // ---------------------------------------------------------------------------
  server.registerTool(
    "set_cell",
    {
      title: "Set NExS Cell Value",
      description:
        "Writes a value to an editable cell in the NExS spreadsheet and returns " +
        "all cells that changed as a result of the backend recalculation. " +
        "Only cells marked as editable in the spreadsheet can be written.",
      inputSchema: {
        cell_ref: z
          .string()
          .describe("Cell address such as 'A1', 'B17', or 'Sheet1!A1'."),
        value: z
          .union([z.string(), z.number()])
          .describe("New value to write to the cell."),
        sheet: z
          .string()
          .optional()
          .describe(
            "Sheet name. Optional when cell_ref includes the sheet or the spreadsheet has one sheet."
          ),
      },
      outputSchema: {
        revision: z.number().describe("New revision number after the change."),
        changed: z
          .array(
            z.object({
              sheet: z.string(),
              addr: z.string(),
              value: z.union([z.string(), z.number()]),
              text: z.string(),
              datatype: z.enum(["numeric", "string", "error", "n/a"]),
            })
          )
          .describe("Cells that changed as a result of this write."),
      },
    },
    async ({ cell_ref, value, sheet }): Promise<CallToolResult> => {
      if (!nexsSession) {
        return {
          isError: true,
          content: [
            { type: "text", text: "No NExS spreadsheet is loaded. Call render_nexs_spreadsheet first." },
          ],
        };
      }

      const { sheet: parsedSheet, cell: cellAddr } = parseCellRef(cell_ref);
      const sheetName =
        sheet ??
        parsedSheet ??
        nexsSession.views.find((v) => !v.isInvisible)?.sheetName ??
        null;

      if (!sheetName) {
        return {
          isError: true,
          content: [
            {
              type: "text",
              text:
                "Cannot determine sheet name. Provide 'sheet' or use 'Sheet1!A1' notation.",
            },
          ],
        };
      }

      let result: NexsInteractResult;
      try {
        result = await nexsInteract(
          nexsSession.appUuid,
          nexsSession.sessionId,
          nexsSession.revision,
          [[sheetName, cellAddr, value]],
        );
      } catch (err) {
        return {
          isError: true,
          content: [
            {
              type: "text",
              text: `Failed to write to NExS: ${err instanceof Error ? err.message : String(err)}`,
            },
          ],
        };
      }

      applyDelta(nexsSession, result);

      const changed = result.values.map(([sn, ci]) => ({
        sheet: sn,
        addr: ci.addr,
        value: ci.data,
        text: ci.text,
        datatype: ci.datatype as "numeric" | "string" | "error" | "n/a",
      }));

      const summary =
        changed.length > 0
          ? changed.map((c) => `${c.sheet}!${c.addr} = ${c.text}`).join(", ")
          : `${sheetName}!${cellAddr} set (no downstream changes reported)`;

      return {
        content: [{ type: "text", text: `Set ${sheetName}!${cellAddr} = ${value}. Changes: ${summary}` }],
        structuredContent: { revision: result.revision, changed },
      };
    }
  );

  // ---------------------------------------------------------------------------
  // UI resource
  // ---------------------------------------------------------------------------
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
                csp: { frameDomains: ["https://platform.nexs.com"] },
              },
            },
          },
        ],
      };
    }
  );

  return server;
}
