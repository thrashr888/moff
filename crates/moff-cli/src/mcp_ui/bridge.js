// Minimal MCP Apps view runtime (SEP-1865): JSON-RPC over postMessage with the host.
// Calls render(result, host) for each tool result; the host sends the tool's CallToolResult.
(function () {
  let nextId = 1;
  const pending = new Map();
  const send = (message) => window.parent.postMessage({ jsonrpc: "2.0", ...message }, "*");
  const request = (method, params) =>
    new Promise((resolve, reject) => {
      const id = nextId++;
      pending.set(id, { resolve, reject });
      send({ id, method, params });
    });
  const notify = (method, params) => send({ method, params });
  let host = { theme: matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light", locale: navigator.language };
  // A host's style variables (SEP-1865 hostContext.styles) win over the card's own palette, so a
  // host can tint cards to its theme. Moff only knows light and dark itself.
  const VARIABLES = {
    "--color-text-primary": "--fg", "--color-text-secondary": "--muted", "--color-background-secondary": "--card",
    "--color-border-primary": "--line", "--color-text-info": "--accent", "--font-sans": "font-family",
  };
  const applied = new Set();
  const applyTheme = () => {
    const root = document.documentElement;
    root.dataset.theme = host.theme === "dark" ? "dark" : "light";
    // Drop the last context's overrides so a removed one falls back to the theme.
    for (const name of applied) root.style.removeProperty(name);
    applied.clear();
    for (const [name, value] of Object.entries(host.styles?.variables ?? {})) {
      if (typeof value !== "string" || !value) continue;
      const target = /^--(fg|muted|card|line|accent)$/.test(name) ? name : VARIABLES[name];
      if (target) { root.style.setProperty(target, value); applied.add(target); }
    }
  };

  window.moff = {
    openLink: (url) => request("ui/open-link", { url }).catch(() => {}),
    host: () => host,
  };

  // Measure the body: the root element is never shorter than the frame, so it can't shrink.
  const reportSize = () =>
    notify("ui/notifications/size-changed", { width: document.documentElement.scrollWidth, height: Math.ceil(document.body.getBoundingClientRect().height) });
  new ResizeObserver(reportSize).observe(document.body);

  window.addEventListener("message", (event) => {
    const message = event.data;
    if (!message || message.jsonrpc !== "2.0") return;
    if (message.id != null && pending.has(message.id)) {
      const { resolve, reject } = pending.get(message.id);
      pending.delete(message.id);
      message.error ? reject(message.error) : resolve(message.result);
      return;
    }
    if (message.method === "ui/notifications/tool-result") {
      window.render?.(message.params ?? {}, host);
      reportSize();
    } else if (message.method === "ui/notifications/host-context-changed") {
      host = { ...host, ...(message.params ?? {}) };
      applyTheme();
    }
  });

  // Links inside a card open through the host (in the user's browser), never in the frame.
  document.addEventListener("click", (event) => {
    const link = event.target.closest?.("a[href]");
    if (!link) return;
    event.preventDefault();
    window.moff.openLink(link.href);
  });

  applyTheme();
  request("ui/initialize", {
    protocolVersion: "2026-01-26",
    appInfo: { name: "moff", version: "1" },
    appCapabilities: { availableDisplayModes: ["inline"] },
  })
    .then((result) => {
      host = { ...host, ...(result?.hostContext ?? {}) };
      applyTheme();
    })
    .catch(() => {})
    .finally(() => notify("ui/notifications/initialized", {}));
})();
