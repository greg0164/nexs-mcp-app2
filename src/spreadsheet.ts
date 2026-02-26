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
 * to fetch the last-rendered URL from server memory. This works because the server
 * process persists between requests and stores the last URL in a module-level variable.
 *
 * NExS session capture:
 * After the NExS iframe initialises, its JavaScript calls POST /api/app/{uuid}/init and
 * then sends the init response (including the session UUID and revision) to the parent
 * window via postMessage. We listen for that message and relay the session ID to the
 * server via the internal set_nexs_session tool. This ensures that get_cell and set_cell
 * on the server operate on the same interactive instance that the iframe is displaying.
 */
import {
  App,
  applyDocumentTheme,
  applyHostFonts,
  applyHostStyleVariables,
  type McpUiHostContext,
} from "@modelcontextprotocol/ext-apps";

const root = document.getElementById("app-root")!;

// --- Helper: apply host theme/styles to the document ---
function applyHostContext(ctx: Partial<McpUiHostContext>) {
  if (ctx.theme) applyDocumentTheme(ctx.theme);
  if (ctx.styles?.variables) applyHostStyleVariables(ctx.styles.variables);
  if (ctx.styles?.css?.fonts) applyHostFonts(ctx.styles.css.fonts);
  if (ctx.safeAreaInsets) {
    const { top, right, bottom, left } = ctx.safeAreaInsets;
    root.style.padding = `${top}px ${right}px ${bottom}px ${left}px`;
  }
}

// --- Helper: render an error message in the view ---
function showError(message: string) {
  root.innerHTML = `
    <div class="error">
      <strong>Unable to load spreadsheet</strong>
      <span>${message}</span>
    </div>`;
}

// 1. Create the app instance
// autoResize is true by default — it watches document.body via ResizeObserver
// and reports the size to the host. The min-height on body (in the HTML) ensures
// the host sees a meaningful intrinsic height rather than near-zero.
const app = new App({ name: "NExS Spreadsheet Viewer", version: "1.0.0" });

// URL captured from tool input — used as a reliable fallback in ontoolresult.
// ontoolinput always fires with the raw tool arguments before the result arrives,
// regardless of whether the host forwards structuredContent.
let capturedUrl: string | null = null;

// True once an iframe has been successfully mounted. Prevents the refresh-recovery
// path from overwriting a live ontoolresult that already ran.
let mounted = false;

// How long to wait for ontoolresult before assuming this is a page refresh.
const REFRESH_DELAY_MS = 2000;

// ---------------------------------------------------------------------------
// NExS session capture via postMessage
//
// After initialising, the NExS iframe JavaScript posts its init result
// (session UUID, revision, values, views) to the parent window. We listen
// for that message and relay it to the server so that get_cell / set_cell
// operate on the same interactive session that the iframe is displaying.
// ---------------------------------------------------------------------------

// UUID regex — matches 8-4-4-4-12 hex format.
const UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

function onNexsMessage(event: MessageEvent) {
  if (event.origin !== "https://platform.nexs.com") return;

  const data: unknown = event.data;
  if (!data || typeof data !== "object" || Array.isArray(data)) return;

  const d = data as Record<string, unknown>;

  // NExS sends the init response to the parent after initialising.
  // Try common field names for the session UUID.
  const sessionId =
    typeof d["session"] === "string"
      ? d["session"]
      : typeof d["sessionId"] === "string"
      ? d["sessionId"]
      : typeof d["session_id"] === "string"
      ? d["session_id"]
      : null;

  if (!sessionId || !UUID_RE.test(sessionId)) return;

  const revision = typeof d["revision"] === "number" ? d["revision"] : 0;
  const values = Array.isArray(d["values"]) ? d["values"] : [];
  const views = Array.isArray(d["views"]) ? d["views"] : [];

  // Relay to the server — non-fatal if the tool isn't available.
  app
    .callServerTool({
      name: "set_nexs_session",
      arguments: { session_id: sessionId, revision, values, views },
    })
    .catch(() => {});
}

// Register once at module load so we catch messages even if the iframe
// loads before ontoolresult fires.
window.addEventListener("message", onNexsMessage);

// --- Helper: validate, sanitise, and mount the NExS iframe ---
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

  // Explicitly signal the desired size now that we have content.
  app.sendSizeChanged({ width: 900, height: 550 }).catch(() => {});
}

// 2. Register ALL handlers BEFORE connecting
app.onhostcontextchanged = applyHostContext;

app.onteardown = async () => {
  return {};
};

app.onerror = (error) => {
  console.error("[NExS MCP App] error:", error);
};

// Capture the URL from tool input arguments — these are always forwarded by
// the host, making this the most reliable source for the spreadsheet URL.
app.ontoolinput = (params) => {
  const args = params.arguments as { app_url?: string } | undefined;
  if (args?.app_url) capturedUrl = args.app_url;
};

// Mount the iframe when the tool result arrives.
// Prefer structuredContent (present when outputSchema is declared on the tool);
// fall back to the URL captured from ontoolinput.
app.ontoolresult = (result) => {
  const data = result.structuredContent as { app_url?: string } | null;
  const url = data?.app_url ?? capturedUrl;
  capturedUrl = null;

  if (!url) {
    showError("No spreadsheet URL was provided.");
    return;
  }

  mountSpreadsheet(url);
};

// 3. Connect to the host (triggers initial context delivery and queued tool results)
app.connect().then(() => {
  const ctx = app.getHostContext();
  if (ctx) applyHostContext(ctx);

  // Refresh recovery: give ontoolresult REFRESH_DELAY_MS to fire naturally.
  // If it doesn't arrive (page refresh — host won't re-deliver historical
  // notifications), call the server-side restore tool to fetch the last URL.
  // localStorage is blocked in ChatGPT's sandbox, so we use callServerTool
  // which proxies through the host to our MCP server's in-memory store.
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
