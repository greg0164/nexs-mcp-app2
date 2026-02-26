/**
 * NExS Spreadsheet MCP App View
 *
 * Lifecycle:
 * 1. Register all handlers before calling app.connect()
 * 2. On ontoolinput — capture the app_url from the tool arguments (always forwarded).
 * 3. On ontoolresult — mount the NExS iframe; set mounted=true to cancel the
 *    refresh-recovery timer.
 * 4. On onhostcontextchanged — apply host theme, fonts, and safe-area insets.
 *
 * Refresh recovery:
 * The host does not re-deliver ontoolinput/ontoolresult for historical tool calls
 * when the conversation is revisited (e.g. page refresh). After connect(), we wait
 * REFRESH_DELAY ms for ontoolresult to fire naturally. If it doesn't arrive (refresh
 * case), we restore the last-mounted URL from localStorage. This avoids racing with
 * the live ontoolresult on a fresh first load.
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

// True once an iframe has been successfully mounted (by ontoolresult or
// by the refresh-recovery timer). Used to prevent double-mounting.
let mounted = false;

// localStorage key for refresh recovery.
const STORAGE_KEY = "nexs:spreadsheet:url";

// How long to wait for ontoolresult before assuming this is a page refresh
// and restoring from localStorage. Should be longer than the typical
// server round-trip but short enough to feel responsive on refresh.
const REFRESH_DELAY_MS = 2000;

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

  // Persist for refresh recovery. Wrapped in try/catch because localStorage
  // may be unavailable in some sandbox configurations.
  try {
    localStorage.setItem(STORAGE_KEY, safeUrl);
  } catch {
    // not available — refresh recovery won't work, but everything else will
  }
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

  // Refresh recovery: wait for ontoolresult to fire naturally (fresh tool call).
  // If it hasn't arrived after REFRESH_DELAY_MS, assume this is a page refresh
  // where the host won't re-deliver historical notifications, and restore the
  // last-known URL from localStorage instead.
  setTimeout(() => {
    if (mounted) return; // ontoolresult already handled it
    try {
      const saved = localStorage.getItem(STORAGE_KEY);
      if (saved) mountSpreadsheet(saved);
    } catch {
      // localStorage not available — no refresh recovery
    }
  }, REFRESH_DELAY_MS);
});
