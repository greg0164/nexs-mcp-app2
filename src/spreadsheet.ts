/**
 * NExS Spreadsheet MCP App View
 *
 * Lifecycle:
 * 1. Register all handlers before calling app.connect()
 * 2. On ontoolinput — capture the app_url from the tool arguments (always forwarded).
 * 3. On ontoolresult — mount the NExS iframe using the URL captured from input,
 *    falling back to structuredContent if available.
 * 4. On onhostcontextchanged — apply host theme, fonts, and safe-area insets.
 *
 * Refresh recovery:
 * The host does not re-deliver ontoolinput/ontoolresult for historical tool calls
 * when the conversation is revisited (e.g. page refresh). To handle this, the last
 * successfully mounted URL is persisted in localStorage and restored immediately
 * on connect(). If the host then fires ontoolresult (new call), it overwrites.
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

// localStorage key used to persist the URL across page refreshes.
// The host does not re-deliver tool notifications on revisit, so we store
// the last-mounted URL ourselves and restore it on connect().
const STORAGE_KEY = "nexs:spreadsheet:url";

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
app.connect().then(async () => {
  const ctx = app.getHostContext();
  if (ctx) applyHostContext(ctx);

  // Refresh recovery: restore the last-known URL from localStorage.
  // On a page refresh the host re-mounts the App View but does not re-fire
  // ontoolinput or ontoolresult for historical calls, so the app would
  // otherwise stay on the "Loading…" screen indefinitely.
  // If ontoolresult fires afterwards (fresh tool call) it will overwrite.
  try {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (saved) mountSpreadsheet(saved);
  } catch {
    // localStorage not available in this sandbox — no refresh recovery
  }

  // Try fullscreen first; if the host doesn't support it, the catch is silent.
  if (ctx?.availableDisplayModes?.includes("fullscreen")) {
    await app.requestDisplayMode({ mode: "fullscreen" }).catch(() => {});
  }
});
