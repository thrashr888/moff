async function(resource) {
    if (!["https://outlook.office.com", "https://outlook.cloud.microsoft", "https://outlook.live.com"].includes(location.origin)
        || ![location.origin, "https://graph.microsoft.com", "https://substrate.office.com"].includes(resource)) {
        throw new Error("Unsupported browser authentication resource");
    }
    const chunks = window.rspackChunkOwa || window.webpackChunkOwa;
    if (!chunks) return {status: 425};
    let runtime;
    chunks.push([["__moff_auth_" + crypto.randomUUID()], {}, value => { runtime = value; }]);
    chunks.pop();
    if (!runtime) return {status: 422};
    // Outlook's current Rspack build changed the auth module ID and no longer needs mailbox context.
    const usesLegacyAuthModule = Boolean(runtime.m[121455]);
    const authId = usesLegacyAuthModule ? 121455 : 100148;
    const mailboxId = runtime.m[429896] ? 429896 : null;
    if (!runtime.m[authId] || (usesLegacyAuthModule && mailboxId === null)) return {status: 422};
    const auth = runtime(authId), mailbox = mailboxId === null ? null : runtime(mailboxId);
    if (typeof auth.ij !== "function" || (mailbox && typeof mailbox.A !== "function")) return {status: 422};
    if (location.origin === "https://outlook.live.com" && resource === location.origin) {
        // Personal accounts call the mail service with an MSA service-scope token, which
        // Outlook returns already formatted as an MSAuth1.0 authorization value.
        if (typeof auth.a9 !== "function" || !mailbox) return {status: 422};
        const value = await auth.a9("service::outlook.office.com::MBI_SSL", undefined, undefined, undefined, mailbox.A());
        if (typeof value !== "string" || !value) return {status: 401};
        const authorization = /^msauth1\.0 /i.test(value) ? value : 'MSAuth1.0 usertoken="' + value + '", type="MSACT"';
        return {status: 200, accessToken: value, authorization};
    }
    const token = await auth.ij("MoffRead", mailbox?.A(), {resource});
    if (!token || !token.AccessToken || token.TokenResultCode !== 0) return {status: 401};
    return {status: 200, accessToken: token.AccessToken, authorization: "Bearer " + token.AccessToken};
}
