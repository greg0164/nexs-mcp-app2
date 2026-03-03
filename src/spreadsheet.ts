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
 * ChatGPT creates a fresh App View instance for every app tool call. Only the
 * render_nexs_spreadsheet App View mounts a NExS iframe. The set_cell App View
 * must NOT mount a second iframe — doing so registers a second React Router
 * navigation blocker in ChatGPT, which crashes the page.
 *
 * Instead, the set_cell App View uses BroadcastChannel("nexs-inputs") to signal
 * the render App View to forward the {op:"input"} postMessage to its existing
 * iframe. If BroadcastChannel is unavailable (sandboxed opaque origin), the
 * set_cell App View falls back to showing a compact text summary of what changed.
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
// BroadcastChannel — cross-App-View input relay
// ---------------------------------------------------------------------------
// The set_cell App View sends {viewIndex, addr, value} here.
// The render App View receives it and forwards to the NExS iframe via postMessage.
// Falls back silently if BroadcastChannel is unavailable (opaque sandbox origin).
// ---------------------------------------------------------------------------
let nexsInputChannel: BroadcastChannel | null = null;
try {
  nexsInputChannel = new BroadcastChannel("nexs-inputs");
} catch {
  // BroadcastChannel not available — cross-App-View update not possible
}

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

// BroadcastChannel listener: forward inputs from set_cell App View to the iframe.
if (nexsInputChannel) {
  nexsInputChannel.onmessage = (e: MessageEvent) => {
    const msg = e.data as { viewIndex?: unknown; addr?: unknown; value?: unknown } | null;
    if (typeof msg?.viewIndex === "number" && typeof msg?.addr === "string") {
      const iframe = root.querySelector("iframe") as HTMLIFrameElement | null;
      console.log("[nexs] BroadcastChannel: forwarding input to iframe:", msg, "iframe=", iframe ? "found" : "null");
      if (iframe?.contentWindow) {
        iframe.contentWindow.postMessage(
          JSON.stringify({ op: "input", id: IFRAME_ID, viewIndex: msg.viewIndex, cell: msg.addr, value: msg.value }),
          NEXS_ORIGIN
        );
      }
    }
  };
}

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

  console.log("[nexs] ontoolresult fired:", { viewIndex: structured?.viewIndex, addr: structured?.addr, app_url: structured?.app_url });

  // set_cell result: viewIndex and addr present in structuredContent.
  if (structured && typeof structured.viewIndex === "number" && structured.addr) {
    const iframe = root.querySelector("iframe") as HTMLIFrameElement | null;
    console.log("[nexs] set_cell path: iframe=", iframe ? "found" : "null");

    if (iframe?.contentWindow) {
      // Fast path: this App View already has a mounted iframe (e.g. same App View
      // instance reused by host). Send input directly via postMessage.
      const msg = JSON.stringify({
        op: "input",
        id: IFRAME_ID,
        viewIndex: structured.viewIndex,
        cell: structured.addr,
        value: structured.value,
      });
      console.log("[nexs] postMessage to iframe:", msg);
      iframe.contentWindow.postMessage(msg, NEXS_ORIGIN);
    } else {
      // No iframe in this App View (fresh set_cell App View created by ChatGPT).
      // Do NOT mount a second NExS iframe — it causes React Router to register a
      // second navigation blocker in ChatGPT, crashing the page with a 404 loop.
      //
      // Instead, broadcast the input to the render App View via BroadcastChannel.
      // If BroadcastChannel is unavailable (sandboxed opaque origin), fall back to
      // a compact text summary (the model response already lists the changes).
      console.log("[nexs] set_cell: broadcasting input via BroadcastChannel");
      if (nexsInputChannel) {
        try {
          nexsInputChannel.postMessage({
            viewIndex: structured.viewIndex as number,
            addr: structured.addr as string,
            value: structured.value,
          });
        } catch {
          // ignore — BroadcastChannel may be closed
        }
      }

      // Show compact confirmation in this App View slot.
      const changed = (structured.changed as Array<{ addr: string; text: string }> | undefined) ?? [];
      const summary = changed.length > 0
        ? changed.map((c) => `${c.addr} → ${c.text}`).join(" · ")
        : `${String(structured.addr)} = ${String(structured.value)}`;
      root.innerHTML = `<div style="padding:10px 14px;font:13px/1.6 system-ui,sans-serif;color:#555;background:#f9fafb;border-radius:6px">
        <strong style="color:#222">Updated:</strong> ${summary}
      </div>`;
      app.sendSizeChanged({ width: 900, height: 46 }).catch(() => {});
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
