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
// We relay initApp and updateCellMap to the server's update_nexs_cells tool.
// initApp may carry the iframe's session UUID and revision; if so, the server
// adopts them so set_cell interact calls target the same live session.
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

  if (data.op === "initApp" && Array.isArray(data.views)) {
    // Extract per-view cell maps (same shape as updateCellMap.cells).
    const views = data.views as Array<{ cells?: Record<string, unknown>; sheetName?: string }>;
    const cells = views.map((v) => v.cells ?? {});
    // Also pass the sheetName for each view — the server uses this when its own
    // views list is empty (e.g. because nexsInit failed) so cells are not lost.
    const sheetNames = views.map((v) => v.sheetName ?? null);
    // Pass session/revision if the iframe included them — the server will
    // adopt the iframe's real session UUID so interact calls match.
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

  // Step 3: send the init handshake to the iframe once it has loaded.
  // This triggers the iframe to respond with {op:"initApp",...} which seeds
  // the server's cell cache with the real current session values.
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

  // set_cell result: the server computed the new values and returned viewIndex,
  // addr, and value so we can forward the same input to the NExS iframe via the
  // embed protocol.  This keeps the live iframe in sync with the AI's write.
  if (structured && typeof structured.viewIndex === "number" && structured.addr) {
    const iframe = root.querySelector("iframe") as HTMLIFrameElement | null;
    if (iframe?.contentWindow) {
      iframe.contentWindow.postMessage(
        JSON.stringify({
          op: "input",
          viewIndex: structured.viewIndex,
          cell: structured.addr,
          value: structured.value,
        }),
        "https://platform.nexs.com"
      );
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
  // The model sometimes re-calls render_nexs_spreadsheet; remounting would
  // destroy the iframe, reset the session, and lose user edits.
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

  // ---------------------------------------------------------------------------
  // Polling loop: forward pending set_cell inputs to the NExS iframe.
  //
  // set_cell runs server-side (nexsInteract) but the iframe display only
  // updates when it receives an {op:"input"} postMessage.  The server queues
  // each set_cell input in pendingDisplayInputs; we drain that queue every
  // second and forward each entry to the iframe.  This is the reliable path —
  // it works even when ontoolresult doesn't fire for non-render tools.
  // ---------------------------------------------------------------------------
  setInterval(() => {
    const iframe = root.querySelector("iframe") as HTMLIFrameElement | null;
    if (!iframe?.contentWindow) return;
    app
      .callServerTool({ name: "pop_nexs_display_inputs", arguments: {} })
      .then((result) => {
        const structured = result.structuredContent as {
          inputs?: Array<{ viewIndex: number; addr: string; value: unknown }>;
        } | null;
        for (const input of structured?.inputs ?? []) {
          iframe.contentWindow!.postMessage(
            JSON.stringify({
              op: "input",
              viewIndex: input.viewIndex,
              cell: input.addr,
              value: input.value,
            }),
            NEXS_ORIGIN
          );
        }
      })
      .catch(() => {});
  }, 1000);

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
