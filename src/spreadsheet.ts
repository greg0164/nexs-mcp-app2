/**
 * NExS Spreadsheet MCP App View
 *
 * Lifecycle:
 * 1. Register all handlers before calling app.connect()
 * 2. On ontoolresult — extract the `app_url` from structuredContent and
 *    mount the NExS iframe, replacing the loading placeholder.
 * 3. On onhostcontextchanged — apply host theme, fonts, and safe-area insets.
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
const app = new App({ name: "NExS Spreadsheet Viewer", version: "1.0.0" });

// 2. Register ALL handlers BEFORE connecting
app.onhostcontextchanged = applyHostContext;

app.onteardown = async () => {
  return {};
};

app.onerror = (error) => {
  console.error("[NExS MCP App] error:", error);
};

// Primary handler: the host pushes the tool result after calling render_nexs_spreadsheet.
// Extract the URL from structuredContent and mount the iframe.
app.ontoolresult = (result) => {
  const data = result.structuredContent as { app_url?: string } | null;
  const url = data?.app_url;

  if (!url) {
    showError("No spreadsheet URL was provided.");
    return;
  }

  if (!url.startsWith("https://platform.nexs.com/")) {
    showError(
      "Only URLs from https://platform.nexs.com are supported. " +
        `Received: ${url}`
    );
    return;
  }

  // Sanitise the URL through the URL constructor before injecting into HTML
  let safeUrl: string;
  try {
    safeUrl = new URL(url).toString();
  } catch {
    showError("The provided URL is not valid.");
    return;
  }

  root.innerHTML = `<iframe src="${safeUrl}" allowfullscreen></iframe>`;
};

// 3. Connect to the host (triggers initial context delivery and queued tool results)
app.connect().then(() => {
  const ctx = app.getHostContext();
  if (ctx) applyHostContext(ctx);
});
