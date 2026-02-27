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

/** [sheetname, cellinfo] tuple — the format used by both init and interact. */
type NexsCellEntry = [string, NexsCellInfo];

/**
 * Tracks the active NExS session.
 *
 * Initially populated by a server-side init call in render_nexs_spreadsheet,
 * then updated by set_nexs_session when the App View captures the browser
 * iframe's real session UUID via postMessage.
 *
 * cellCache stores the full current state of all known cells so that get_cell
 * can return values even for cells that haven't changed since the last interact
 * call (interact only returns deltas since the given revision).
 */
interface NexsSession {
  appUuid: string;
  /** Session UUID — starts as the server's own, replaced by the iframe's when available. */
  sessionId: string;
  /** Revision ID of the last interact call; used as the baseline for the next one. */
  revision: number;
  views: NexsView[];
  /**
   * Full cell-value cache keyed by "SheetName!ADDR" (addr uppercased).
   * Populated from init values and kept current by merging interact deltas.
   */
  cellCache: Map<string, { sheetName: string; ci: NexsCellInfo }>;
  /** True once the App View has relayed the iframe's real session UUID. */
  fromBrowser: boolean;
}

// ---------------------------------------------------------------------------
// Module-level store — one process, survives across per-request McpServer
// instances (the module is loaded once per process).
// ---------------------------------------------------------------------------

/** Last-rendered spreadsheet URL — used by the App View's refresh recovery. */
let lastSpreadsheetUrl: string | null = null;

/**
 * Active NExS session.
 * Null until render_nexs_spreadsheet has been called at least once.
 */
let nexsSession: NexsSession | null = null;

// ---------------------------------------------------------------------------
// NExS interact API helpers
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
 * Creates a new NExS session and returns all current cell values.
 * The endpoint is csrf_exempt and works without cookies for public apps.
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
 * Operates on an existing NExS session. Returns only cells that changed
 * since the given revision. DRF's APIView wraps all views with csrf_exempt;
 * CSRF is only enforced by SessionAuthentication when a Django session cookie
 * is present, so unauthenticated server-side requests need no CSRF token.
 *
 * @param inputs  Pass [] for a read-only delta poll (keepalive);
 *                pass [[sheetname, celladdr, value], ...] to write cells.
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

/** Merge interact delta values into the session's cell cache and advance the revision. */
function applyDelta(session: NexsSession, result: NexsInteractResult): void {
  for (const [sn, ci] of result.values) {
    session.cellCache.set(`${sn}!${ci.addr.toUpperCase()}`, { sheetName: sn, ci });
  }
  session.revision = result.revision;
}

/** Build a fresh cell cache from a set of init values. */
function buildCellCache(values: NexsCellEntry[]): NexsSession["cellCache"] {
  const cache: NexsSession["cellCache"] = new Map();
  for (const [sn, ci] of values) {
    cache.set(`${sn}!${ci.addr.toUpperCase()}`, { sheetName: sn, ci });
  }
  return cache;
}

// ---------------------------------------------------------------------------
// MCP server factory
// ---------------------------------------------------------------------------

export function createServer(): McpServer {
  const server = new McpServer({
    name: "NExS Spreadsheet Viewer",
    version: "1.0.0",
  });

  // ---------------------------------------------------------------------------
  // render_nexs_spreadsheet
  // Renders the spreadsheet in-chat and seeds a server-side NExS session so
  // that get_cell / set_cell have a ready baseline before the browser iframe
  // sends its own session via postMessage.
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
          .describe("A published NExS spreadsheet URL (https://platform.nexs.com/...)."),
      },
      outputSchema: {
        app_url: z.string().url().describe("The NExS spreadsheet URL being rendered."),
      },
      _meta: { ui: { resourceUri: RESOURCE_URI } },
    },
    async ({ app_url }): Promise<CallToolResult> => {
      lastSpreadsheetUrl = app_url;

      // Seed a server-side session so get_cell / set_cell work immediately.
      // This session will be superseded by the iframe's real session once
      // set_nexs_session is called from the App View (postMessage capture).
      const appUuid = extractNexsUuid(app_url);
      if (appUuid) {
        nexsInit(appUuid)
          .then((init) => {
            if (nexsSession?.fromBrowser) {
              // Browser session already captured — don't replace it, but
              // backfill any cells that the postMessage didn't include.
              for (const [sn, ci] of init.values) {
                const key = `${sn}!${ci.addr.toUpperCase()}`;
                if (!nexsSession!.cellCache.has(key)) {
                  nexsSession!.cellCache.set(key, { sheetName: sn, ci });
                }
              }
            } else {
              // postMessage hasn't arrived yet — use the server's own session
              // as a fallback until set_nexs_session fires.
              nexsSession = {
                appUuid,
                sessionId: init.sessionId,
                revision: init.revision,
                views: init.views,
                cellCache: buildCellCache(init.values),
                fromBrowser: false,
              };
            }
          })
          .catch((err: unknown) => {
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
  // set_nexs_session — internal, App View only.
  //
  // Called by spreadsheet.ts when it receives the postMessage that NExS sends
  // to the parent window after the iframe's own init completes. This gives the
  // server the session UUID the browser is actually using, so that subsequent
  // get_cell / set_cell calls operate on the same interactive instance.
  // ---------------------------------------------------------------------------
  registerAppTool(
    server,
    "set_nexs_session",
    {
      description:
        "Receives the NExS session ID from the browser iframe after it initialises. " +
        "Internal use only — called by the App View, not the model.",
      inputSchema: {
        session_id: z.string().describe("NExS session UUID from the browser iframe."),
        revision: z.number().int().describe("Current revision number from the iframe's init."),
        values: z
          .array(z.unknown())
          .describe("Initial cell values from the iframe's init response."),
        views: z
          .array(z.unknown())
          .describe("View definitions from the iframe's init response."),
      },
      _meta: { ui: { resourceUri: RESOURCE_URI, visibility: ["app"] } },
    },
    async ({ session_id, revision, values, views }): Promise<CallToolResult> => {
      // Resolve the app UUID: prefer the already-stored session, then fall
      // back to the last-rendered URL. This handles the race where the iframe
      // postMessage arrives before the background nexsInit() call completes
      // (nexsSession is still null at that point).
      const appUuid =
        nexsSession?.appUuid ??
        (lastSpreadsheetUrl ? extractNexsUuid(lastSpreadsheetUrl) : null);

      if (!appUuid) return { content: [] }; // render_nexs_spreadsheet not called yet

      const parsedViews = Array.isArray(views) ? (views as NexsView[]) : [];
      const parsedValues = Array.isArray(values) ? (values as NexsCellEntry[]) : [];

      if (nexsSession) {
        // Upgrade the existing (possibly server-side fallback) session to the
        // browser's real session UUID and revision.
        nexsSession.sessionId = session_id;
        nexsSession.revision = revision;
        nexsSession.fromBrowser = true;
        if (parsedViews.length > 0) nexsSession.views = parsedViews;
      } else {
        // postMessage arrived before the background init completed — create the
        // session now using the browser's data so we don't miss it.
        nexsSession = {
          appUuid,
          sessionId: session_id,
          revision,
          views: parsedViews,
          cellCache: new Map(),
          fromBrowser: true,
        };
      }

      // Merge the iframe's initial cell values (authoritative — they come from
      // the session the user is actually looking at).
      for (const entry of parsedValues) {
        if (Array.isArray(entry) && entry.length === 2) {
          const [sn, ci] = entry as [string, NexsCellInfo];
          if (sn && ci?.addr) {
            nexsSession.cellCache.set(
              `${sn}!${ci.addr.toUpperCase()}`,
              { sheetName: sn, ci }
            );
          }
        }
      }

      return { content: [] };
    }
  );

  // ---------------------------------------------------------------------------
  // get_cell — read a cell from the active session.
  //
  // Calls interact (not init) so it operates on the existing session — the
  // same one the browser iframe is using after set_nexs_session has fired.
  // interact returns only deltas since the last revision; those are merged into
  // the cellCache so we always have the full current state available.
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

      // Sync with the server — fetch any cells that changed since the last
      // revision.  This is what makes get_cell reflect edits the user made
      // in the browser after the initial session was established.
      try {
        const delta = await nexsInteract(
          nexsSession.appUuid,
          nexsSession.sessionId,
          nexsSession.revision,
          [],
        );
        applyDelta(nexsSession, delta);
      } catch (interactErr) {
        // Interact failed (session expired or network error).
        // If we have a browser session, we can't recover without re-capture
        // from the App View.  If this is a server-side fallback session, try
        // re-initing so at least the base values are current.
        if (!nexsSession.fromBrowser) {
          try {
            const init = await nexsInit(nexsSession.appUuid);
            nexsSession.sessionId = init.sessionId;
            nexsSession.revision = init.revision;
            nexsSession.views = init.views;
            nexsSession.cellCache = buildCellCache(init.values);
          } catch {
            // Re-init also failed — proceed with stale cache as last resort.
          }
        } else {
          // Browser session expired mid-conversation.  The App View will send
          // a new set_nexs_session when the iframe reloads.  Report the error
          // rather than silently returning stale data.
          return {
            isError: true,
            content: [
              {
                type: "text",
                text: `NExS session expired (${
                  interactErr instanceof Error ? interactErr.message : String(interactErr)
                }). Please reload the spreadsheet view.`,
              },
            ],
          };
        }
      }

      // Search the cache.
      const cacheKey = sheetName
        ? `${sheetName}!${cellAddr.toUpperCase()}`
        : null;

      let found: { sheetName: string; ci: NexsCellInfo } | undefined;

      if (cacheKey) {
        found = nexsSession.cellCache.get(cacheKey);
      } else {
        // No sheet specified — search all sheets.
        const upper = cellAddr.toUpperCase();
        for (const [key, val] of nexsSession.cellCache) {
          if (key.endsWith(`!${upper}`)) {
            found = val;
            break;
          }
        }
      }

      if (!found) {
        const sheets = [...new Set(
          [...nexsSession.cellCache.values()].map((v) => v.sheetName)
        )];
        return {
          isError: true,
          content: [
            {
              type: "text",
              text:
                `Cell '${cell_ref}' not found in the spreadsheet. ` +
                `Known sheets: ${sheets.join(", ")}`,
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
  // set_cell — write a value to an editable cell in the active session.
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
          .describe("Cell reference such as 'A1', 'B17', or 'Sheet1!A1'."),
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
          .describe(
            "Cells that changed as a result of this input, including the written cell."
          ),
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
                "Cannot determine sheet name. Use the 'sheet' parameter or " +
                "include it in the cell reference (e.g. 'Sheet1!A1').",
            },
          ],
        };
      }

      const doInteract = (sess: NexsSession) =>
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
          // Note: after re-init the session is a fresh server-side one, not the
          // browser's. fromBrowser is reset to false so set_nexs_session can
          // upgrade it again when the next postMessage arrives.
          nexsSession = {
            appUuid: nexsSession.appUuid,
            sessionId: init.sessionId,
            revision: init.revision,
            views: init.views,
            cellCache: buildCellCache(init.values),
            fromBrowser: false,
          };
          result = await doInteract(nexsSession);
        } catch (retryErr) {
          return {
            isError: true,
            content: [
              {
                type: "text",
                text: `Failed to write to NExS spreadsheet: ${
                  retryErr instanceof Error ? retryErr.message : String(retryErr)
                }`,
              },
            ],
          };
        }
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
