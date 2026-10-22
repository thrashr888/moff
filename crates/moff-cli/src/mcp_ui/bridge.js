// Minimal MCP Apps view runtime (SEP-1865): JSON-RPC over postMessage with the host.
// Calls render(result, host) for each tool result; the host sends the tool's CallToolResult.
(function () {
  let nextId = 1;
  const pending = new Map();
  const send = (message) => window.parent.postMessage({ jsonrpc: "2.0", ...message }, "*");
  const request = (method, params) =>
    new Promise((resolve, reject) => {
      const id = nextId++;
      const timer = setTimeout(() => {
        pending.delete(id);
        reject(new Error("The host did not confirm a result. Check the mailbox before retrying an action."));
      }, 120000);
      pending.set(id, {
        resolve: value => { clearTimeout(timer); resolve(value); },
        reject: error => { clearTimeout(timer); reject(error); },
      });
      send({ id, method, params });
    });
  const notify = (method, params) => send({ method, params });
  let host = { theme: matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light", locale: navigator.language };
  const applyTheme = () => document.documentElement.dataset.theme = host.theme || "light";
  const reportError = (error) => {
    let status=document.getElementById("moff-host-error");
    if (!status) {
      status=document.createElement("p");
      status.id="moff-host-error";
      status.setAttribute("role","alert");
      document.body.appendChild(status);
    }
    status.textContent=error?.message || "The host could not complete the request.";
  };

  window.moff = {
    openLink: (url) => request("ui/open-link", { url }),
    callTool: (name, args) => request("tools/call", {name, arguments: args}),
    host: () => host,
  };

  // Measure the body: the root element is never shorter than the frame, so it can't shrink.
  const reportSize = () =>
    notify("ui/notifications/size-changed", { width: document.documentElement.scrollWidth, height: Math.ceil(document.body.getBoundingClientRect().height) });
  new ResizeObserver(reportSize).observe(document.body);

  window.addEventListener("message", (event) => {
    // Tavern installs a parent shim and injects native events without a WindowProxy.
    const nativeHost = event.source === null && window.webkit?.messageHandlers?.mcp;
    if (event.source !== window.parent && !nativeHost) return;
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
    window.moff.openLink(link.href).catch(reportError);
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
    .catch(reportError)
    .finally(() => notify("ui/notifications/initialized", {}));
})();
