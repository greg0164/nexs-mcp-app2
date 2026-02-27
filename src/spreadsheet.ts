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
// between the host page and the NExS iframe.  The iframe sends:
//
//   initApp      — full cell state after the iframe calls init
//   updateCellMap — delta of cells that changed after a user edit
//
// We relay "updateCellMap" to the server-side update_nexs_cells tool so the
// server's cell cache stays in sync with what the user sees in the live iframe.
// This is what makes get_cell accurate after user edits.
window.addEventListener("message", (e) => {
  if (e.data === "hello") return;
  let data: Record<string, unknown>;
  try {
    data = JSON.parse(e.data as string) as Record<string, unknown>;
  } catch {
    return;
  }
  // Only handle messages from the NExS app origin.
  if (!e.origin.startsWith("https://platform.nexs.com")) return;
  if (data.op === "updateCellMap" && Array.isArray(data.cells)) {
    app
      .callServerTool({ name: "update_nexs_cells", arguments: { cells: data.cells } })
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

  root.innerHTML = `<iframe src="${safeUrl}" allowfullscreen></iframe>`;
  mounted = true;

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
