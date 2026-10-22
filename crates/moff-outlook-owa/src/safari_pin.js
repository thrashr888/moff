async function(config) {
    if (!["https://outlook.office.com", "https://outlook.cloud.microsoft"].includes(location.origin)
        || config.allow_mutation !== true || typeof config.id !== "string" || !config.id.trim()
        || typeof config.pinned !== "boolean") throw new Error("Invalid native pin request");
    let runtime;
    const chunks = window.webpackChunkOwa;
    if (!chunks) return {dispatched:false, reason:"outlook_not_ready"};
    chunks.push([["__moff_pin_" + crypto.randomUUID()], {}, value => {runtime=value;}]);
    chunks.pop();
    if (!runtime?.m[708334] || !runtime.m[429896]
        || typeof runtime(708334).pinUnpinItems !== "function"
        || typeof runtime(429896).A !== "function") {
        return {dispatched:false, reason:"auth_adapter_changed"};
    }
    // Observed Outlook PinItem helper; no table argument, so no optimistic list mutation.
    const result = await runtime(708334).pinUnpinItems([config.id], runtime(429896).A(), config.pinned);
    return {dispatched:true, success:result?.success === true};
}
