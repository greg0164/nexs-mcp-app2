/**
 * NExS Spreadsheet MCP App View
 *
 * Lifecycle:
 * 1. Register all handlers before calling app.connect()
 * 2. On ontoolinput — capture the app_url from the tool arguments (always forwarded).
 * 3. On ontoolresult — mount the NExS iframe; set mounted=true.
 * 4. On onhostcontextchanged — apply host theme, fonts, and safe-area insets.
 *
 * Refresh recovery:
 * The host does not re-deliver ontoolinput/ontoolresult for historical tool calls
 * when the conversation is revisited (e.g. page refresh). After connect(), we wait
 * REFRESH_DELAY_MS for ontoolresult to fire naturally. If it doesn't arrive (refresh
 * case), we call the server-side restore_nexs_spreadsheet tool via callServerTool()
 * to fetch the last-rendered URL from server memory.
 *
 * Multi-App-View isolation:
 * ChatGPT creates a fresh App View for every app tool call. The
 * render_nexs_spreadsheet App View is the "primary" instance — it relays
 * NExS iframe messages (initApp, updateCellMap) to the server via
 * callServerTool so get_cell stays in sync with user edits.
 *
 * The set_cell App View is a "display-only" secondary instance. It mounts
 * a fresh NExS iframe which naturally loads the NExS session that was just
 * updated server-side by nexsInteract, showing the AI's changes.
 * Display-only mode suppresses all callServerTool relay calls from the iframe
 * to prevent MCP interference with concurrent model tool calls.
 *
 * Note: {op:"input"} postMessages are silently ignored by the NExS embed.
 * Note: Reloading the render App View's existing iframe resets to published
 *       defaults (NExS session does not survive iframe destruction).
 */
import {
  App,
  applyDocumentTheme,
  applyHostFonts,
  applyHostStyleVariables,
  type McpUiHostContext,
} from "@modelcontextprotocol/ext-apps";

const root = document.getElementById("app-root")!;

function applyHostContext(ctx: Partial<McpUiHostContext>) {
  if (ctx.theme) applyDocumentTheme(ctx.theme);
  if (ctx.styles?.variables) applyHostStyleVariables(ctx.styles.variables);
  if (ctx.styles?.css?.fonts) applyHostFonts(ctx.styles.css.fonts);
  if (ctx.safeAreaInsets) {
    const { top, right, bottom, left } = ctx.safeAreaInsets;
    root.style.padding = `${top}px ${right}px ${bottom}px ${left}px`;
  }
}

function showError(message: string) {
  root.innerHTML = `
    <div class="error">
      <strong>Unable to load spreadsheet</strong>
      <span>${message}</span>
    </div>`;
}

const app = new App({ name: "NExS Spreadsheet Viewer", version: "1.0.0" });

let capturedUrl: string | null = null;
let mounted = false;
const REFRESH_DELAY_MS = 2000;

/**
 * True when this is a display-only App View (set_cell).
 * Suppresses callServerTool relay calls so they don't interfere with
 * concurrent model tool calls on the same MCP session.
 */
let isDisplayOnly = false;

// ---------------------------------------------------------------------------
// NExS embed protocol — postMessage relay
// ---------------------------------------------------------------------------
// The NExS embed protocol (nexs_embed.js) uses bidirectional postMessage
// between the host page and the NExS iframe.
//
// HANDSHAKE (must complete before the iframe sends anything useful):
//   1. Iframe → parent: "hello"   (raw string — origin verification)
//   2. Parent → iframe: "hello"   (echo back)
//   3. Parent → iframe: {op:"init", id:iframeId}  (sent on iframe "load")
//   4. Iframe → parent: {op:"initApp", id, name, views, session?, revision?}
//
// Without step 3 the iframe stays silent — no initApp, no updateCellMap.
//
// ONGOING:
//   Iframe → parent: {op:"updateCellMap", id, cells:[...]}  on every recalc
//
// PRIMARY App Views relay initApp/updateCellMap to the server.
// DISPLAY-ONLY App Views skip the relay entirely.
const IFRAME_ID = "nexs-iframe-0";
const NEXS_ORIGIN = "https://platform.nexs.com";

window.addEventListener("message", (e) => {
  // Step 2: echo "hello" back for the NExS origin-verification handshake.
  if (e.data === "hello") {
    (e.source as Window | null)?.postMessage("hello", e.origin);
    return;
  }

  let data: Record<string, unknown>;
  try {
    data = JSON.parse(e.data as string) as Record<string, unknown>;
  } catch {
    return;
  }
  // Only handle messages from the NExS app origin.
  if (!e.origin.startsWith(NEXS_ORIGIN)) return;

  // Display-only App Views skip relay — would interfere with model tool calls.
  if (isDisplayOnly) return;

  if (data.op === "initApp" && Array.isArray(data.views)) {
    const views = data.views as Array<{ cells?: Record<string, unknown>; sheetName?: string }>;
    const cells = views.map((v) => v.cells ?? {});
    const sheetNames = views.map((v) => v.sheetName ?? null);
    const args: Record<string, unknown> = { cells, sheetNames, isInitApp: true };
    if (typeof data.session === "string") args.sessionId = data.session;
    if (typeof data.revision === "number") args.revision = data.revision;
    app.callServerTool({ name: "update_nexs_cells", arguments: args }).catch(() => {});
  } else if (data.op === "updateCellMap" && Array.isArray(data.cells)) {
    app
      .callServerTool({
        name: "update_nexs_cells",
        arguments: { cells: data.cells, isInitApp: false },
      })
      .catch(() => {});
  }
});

function mountSpreadsheet(url: string) {
  if (!url.startsWith("https://platform.nexs.com/")) {
    showError(
      "Only URLs from https://platform.nexs.com are supported. " +
        `Received: ${url}`
    );
    return;
  }

  let safeUrl: string;
  try {
    safeUrl = new URL(url).toString();
  } catch {
    showError("The provided URL is not valid.");
    return;
  }

  root.innerHTML = `<iframe id="${IFRAME_ID}" src="${safeUrl}" allowfullscreen></iframe>`;
  mounted = true;

  const iframe = root.querySelector("iframe") as HTMLIFrameElement;
  iframe.addEventListener("load", () => {
    iframe.contentWindow?.postMessage(
      JSON.stringify({ op: "init", id: IFRAME_ID }),
      NEXS_ORIGIN
    );
  });

  app.sendSizeChanged({ width: 900, height: 550 }).catch(() => {});
}

app.onhostcontextchanged = applyHostContext;

app.onteardown = async () => {
  return {};
};

app.onerror = (error) => {
  console.error("[NExS MCP App] error:", error);
};

app.ontoolinput = (params) => {
  const args = params.arguments as { app_url?: string } | undefined;
  if (args?.app_url) capturedUrl = args.app_url;
};

app.ontoolresult = (result) => {
  const structured = result.structuredContent as Record<string, unknown> | null;

  console.log("[nexs] ontoolresult fired:", { viewIndex: structured?.viewIndex, addr: structured?.addr, app_url: structured?.app_url });

  // set_cell result: viewIndex and addr present in structuredContent.
  if (structured && typeof structured.viewIndex === "number" && structured.addr) {
    const iframe = root.querySelector("iframe") as HTMLIFrameElement | null;
    console.log("[nexs] set_cell path: iframe=", iframe ? "found" : "null");

    if (!iframe && !mounted) {
      // Fresh set_cell App View. Mount in display-only mode so the iframe
      // naturally loads the NExS session that nexsInteract just updated.
      // Relay is suppressed to avoid MCP interference with model tool calls.
      console.log("[nexs] set_cell: mounting display-only iframe");
      isDisplayOnly = true;
      const url = structured.app_url as string | undefined;
      if (url) mountSpreadsheet(url);
    } else {
      console.log("[nexs] set_cell: iframe exists — changes shown in model response");
    }
    return;
  }

  // render_nexs_spreadsheet result: mount the iframe.
  const url = (structured?.app_url as string | undefined) ?? capturedUrl;
  capturedUrl = null;

  if (!url) {
    showError("No spreadsheet URL was provided.");
    return;
  }

  // Skip remounting if the iframe is already showing this URL.
  const existing = root.querySelector("iframe") as HTMLIFrameElement | null;
  try {
    if (existing && existing.src === new URL(url).toString()) return;
  } catch {
    // URL parse failed — fall through to mountSpreadsheet
  }

  mountSpreadsheet(url);
};

app.connect().then(() => {
  const ctx = app.getHostContext();
  if (ctx) applyHostContext(ctx);

  setTimeout(async () => {
    if (mounted) return;
    try {
      const result = await app.callServerTool({
        name: "restore_nexs_spreadsheet",
        arguments: {},
      });
      const data = result.structuredContent as { app_url?: string | null } | null;
      if (data?.app_url && !mounted) mountSpreadsheet(data.app_url);
    } catch {
      // callServerTool not supported or restore tool unavailable
    }
  }, REFRESH_DELAY_MS);
});
