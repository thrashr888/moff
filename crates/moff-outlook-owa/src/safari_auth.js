async function(resource) {
    if (!["https://outlook.office.com", "https://outlook.cloud.microsoft", "https://outlook.live.com"].includes(location.origin)
        || ![location.origin, "https://graph.microsoft.com", "https://substrate.office.com"].includes(resource)) {
        throw new Error("Unsupported browser authentication resource");
    }
    if (!window.webpackChunkOwa) return {status: 425};
    let runtime;
    const chunks = window.webpackChunkOwa;
    chunks.push([["__moff_auth_" + crypto.randomUUID()], {}, value => { runtime = value; }]);
    chunks.pop();
    if (!runtime || !runtime.m[121455] || !runtime.m[429896]) return {status: 422};
    const auth = runtime(121455), mailbox = runtime(429896);
    if (typeof auth.ij !== "function" || typeof mailbox.A !== "function") return {status: 422};
    const token = await auth.ij("MoffRead", mailbox.A(), {resource});
    if (!token || !token.AccessToken || token.TokenResultCode !== 0) return {status: 401};
    return {status: 200, accessToken: token.AccessToken};
}
