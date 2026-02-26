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

/** [sheetname, cellinfo] tuple — the format used by both init and interact */
type NexsCellEntry = [string, NexsCellInfo];

interface NexsSession {
  appUuid: string;
  sessionId: string;
  revision: number;
  views: NexsView[];
}

// ---------------------------------------------------------------------------
// Module-level store — survives across per-request McpServer instances because
// this module is loaded once per process.
// ---------------------------------------------------------------------------

/** Last-rendered spreadsheet URL — used by the App View's refresh recovery. */
let lastSpreadsheetUrl: string | null = null;

/** Active NExS interact session — used by get_cell / set_cell. */
let nexsSession: NexsSession | null = null;

// ---------------------------------------------------------------------------
// NExS interact API helpers
// ---------------------------------------------------------------------------

/** Extract the NExS app UUID from a platform URL. */
function extractNexsUuid(url: string): string | null {
  const m = url.match(
    /\/([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})/i
  );
  return m ? m[1] : null;
}

/** Parse "Sheet1!A1" → {sheet: "Sheet1", cell: "A1"} or "A1" → {sheet: null, cell: "A1"} */
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
 * Call the NExS init API to start a session and get all current cell values.
 *
 * The init endpoint is csrf_exempt and works without authentication for
 * public (open-mode) NExS apps. It returns a session UUID that can be used
 * for subsequent interact calls.
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
 * Call the NExS interact API to read or write cells.
 *
 * Pass an empty inputs array to perform a keepalive / get-changes-only poll.
 * Pass [[sheetname, celladdr, value], ...] to write cell values.
 *
 * DRF's APIView.as_view() wraps the view with csrf_exempt; CSRF is only
 * enforced by SessionAuthentication when a session cookie is present. For
 * unauthenticated server-side requests (no session cookie), no CSRF is needed.
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

// ---------------------------------------------------------------------------
// MCP server factory
// ---------------------------------------------------------------------------

/**
 * Creates a new MCP server instance with the NExS spreadsheet tool and
 * its corresponding UI resource registered.
 */
export function createServer(): McpServer {
  const server = new McpServer({
    name: "NExS Spreadsheet Viewer",
    version: "1.0.0",
  });

  // ---------------------------------------------------------------------------
  // render_nexs_spreadsheet — renders the spreadsheet in-chat and initialises
  // the interact session so that get_cell / set_cell can be used immediately.
  // ---------------------------------------------------------------------------
  registerAppTool(
    server,
    "render_nexs_spreadsheet",
    {
      title: "Render NExS Spreadsheet",
      description:
        "Renders a live, interactive NExS spreadsheet inline in the conversation. " +
        "Use when the user provides a published NExS platform URL " +
        "(https://platform.nexs.com/...). " +
        "After rendering, use get_cell to read cell values and set_cell to write them.",
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

      // Initialise a NExS interact session in the background so that
      // get_cell / set_cell can work without a cold-start delay.
      const appUuid = extractNexsUuid(app_url);
      if (appUuid) {
        nexsInit(appUuid)
          .then((init) => {
            nexsSession = {
              appUuid,
              sessionId: init.sessionId,
              revision: init.revision,
              views: init.views,
            };
          })
          .catch((err: unknown) => {
            // Non-fatal — get_cell / set_cell will init on demand.
            console.error("[NExS] Background session init failed:", err);
          });
      }

      return {
        content: [{ type: "text", text: `Rendering NExS spreadsheet: ${app_url}` }],
        structuredContent: { app_url },
      };
    }
  );

  // ---------------------------------------------------------------------------
  // restore_nexs_spreadsheet — internal, App View only.
  // Invisible to the LLM. The App calls this on page refresh.
  // ---------------------------------------------------------------------------
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

  // ---------------------------------------------------------------------------
  // get_cell — read a cell value from the displayed NExS spreadsheet.
  // Calls the NExS init API fresh to get the latest cell values.
  // ---------------------------------------------------------------------------
  server.registerTool(
    "get_cell",
    {
      title: "Get NExS Cell Value",
      description:
        "Reads the current value of a cell from the NExS spreadsheet. " +
        "Returns the formatted text, raw data value, and data type. " +
        "Requires render_nexs_spreadsheet to have been called first.",
      inputSchema: {
        cell_ref: z
          .string()
          .describe(
            "Cell reference such as 'A1', 'B17', or 'Sheet1!A1'. " +
            "Named cells defined in the spreadsheet are also supported."
          ),
        sheet: z
          .string()
          .optional()
          .describe(
            "Sheet name to look in. Optional when the cell reference already " +
            "includes the sheet (e.g. 'Sheet1!A1') or the spreadsheet has only one sheet."
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
          content: [
            {
              type: "text",
              text: "No NExS spreadsheet is loaded. Call render_nexs_spreadsheet first.",
            },
          ],
          isError: true,
        };
      }

      const { sheet: parsedSheet, cell: cellAddr } = parseCellRef(cell_ref);
      const sheetName =
        sheet ??
        parsedSheet ??
        nexsSession.views.find((v) => !v.isInvisible)?.sheetName ??
        null;

      let init: NexsInitResult;
      try {
        init = await nexsInit(nexsSession.appUuid);
      } catch (err) {
        return {
          content: [
            {
              type: "text",
              text: `Failed to read from NExS spreadsheet: ${
                err instanceof Error ? err.message : String(err)
              }`,
            },
          ],
          isError: true,
        };
      }

      const entry = init.values.find(([sn, ci]) => {
        const sheetMatch = sheetName ? sn === sheetName : true;
        return sheetMatch && ci.addr.toUpperCase() === cellAddr.toUpperCase();
      });

      if (!entry) {
        const sheets = [...new Set(init.values.map(([sn]) => sn))];
        return {
          content: [
            {
              type: "text",
              text:
                `Cell '${cell_ref}' not found in the spreadsheet. ` +
                `Available sheets: ${sheets.join(", ")}`,
            },
          ],
          isError: true,
        };
      }

      const [foundSheet, ci] = entry;
      return {
        content: [
          {
            type: "text",
            text: `${foundSheet}!${ci.addr} = ${ci.text} (${ci.datatype})`,
          },
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
  // set_cell — write a value to an editable cell in the NExS spreadsheet,
  // triggering the backend recalculation.
  // ---------------------------------------------------------------------------
  server.registerTool(
    "set_cell",
    {
      title: "Set NExS Cell Value",
      description:
        "Writes a value to an editable cell in the NExS spreadsheet and returns " +
        "all cells that changed as a result of the backend recalculation. " +
        "Only cells marked as editable in the spreadsheet can be written. " +
        "Requires render_nexs_spreadsheet to have been called first.",
      inputSchema: {
        cell_ref: z
          .string()
          .describe(
            "Cell reference such as 'A1', 'B17', or 'Sheet1!A1'."
          ),
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
          .describe("Cells that changed as a result of this input, including the written cell."),
      },
    },
    async ({ cell_ref, value, sheet }): Promise<CallToolResult> => {
      if (!nexsSession) {
        return {
          content: [
            {
              type: "text",
              text: "No NExS spreadsheet is loaded. Call render_nexs_spreadsheet first.",
            },
          ],
          isError: true,
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
          content: [
            {
              type: "text",
              text:
                "Cannot determine sheet name. Use the 'sheet' parameter or " +
                "include it in the cell reference (e.g. 'Sheet1!A1').",
            },
          ],
          isError: true,
        };
      }

      const doInteract = async (sess: NexsSession): Promise<NexsInteractResult> =>
        nexsInteract(sess.appUuid, sess.sessionId, sess.revision, [
          [sheetName, cellAddr, value],
        ]);

      let result: NexsInteractResult;
      try {
        result = await doInteract(nexsSession);
      } catch {
        // Session may have expired — re-initialise and retry once.
        try {
          const init = await nexsInit(nexsSession.appUuid);
          nexsSession = {
            appUuid: nexsSession.appUuid,
            sessionId: init.sessionId,
            revision: init.revision,
            views: init.views,
          };
          result = await doInteract(nexsSession);
        } catch (retryErr) {
          return {
            content: [
              {
                type: "text",
                text: `Failed to write to NExS spreadsheet: ${
                  retryErr instanceof Error ? retryErr.message : String(retryErr)
                }`,
              },
            ],
            isError: true,
          };
        }
      }

      // Update the stored revision so future set_cell calls use the right base.
      nexsSession.revision = result.revision;

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
          : `${sheetName}!${cellAddr} set (no downstream changes detected)`;

      return {
        content: [{ type: "text", text: `Set ${sheetName}!${cellAddr} = ${value}. ${summary}` }],
        structuredContent: { revision: result.revision, changed },
      };
    }
  );

  // ---------------------------------------------------------------------------
  // UI resource — single-file bundled HTML View.
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
