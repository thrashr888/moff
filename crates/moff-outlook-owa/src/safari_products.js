async function(config, acquireToken, readResponse) {
    if (!["https://outlook.office.com", "https://outlook.cloud.microsoft"].includes(location.origin)) {
        throw new Error("Product reads require an HTTPS Outlook document");
    }
    const rawPath = config.url.split("?")[0].replace(/^https:\/\/graph\.microsoft\.com/, "");
    if (rawPath.includes("\\") || rawPath.includes("//") || rawPath.endsWith("/")
        || rawPath.split("/").some(segment => /^(?:\.|%2e){1,2}$/i.test(segment))) {
        throw new Error("Unsupported product read URL");
    }
    const url = new URL(config.url);
    const paths = [
        /^\/v1\.0\/me\/(people|manager|directReports)$/,
        /^\/v1\.0\/users\/[^/]+(\/(manager|directReports))?$/,
        /^\/v1\.0\/groups$/,
        /^\/v1\.0\/me\/drive\/(root|items\/[^/]+)\/children$/,
        /^\/v1\.0\/me\/drive\/items\/[^/]+$/,
        /^\/v1\.0\/me\/drive\/root\/search\(q='\.loop'\)$/,
    ];
    if (url.origin !== "https://graph.microsoft.com" || url.username || url.password || url.hash
        || !paths.some(pattern => pattern.test(url.pathname))
        || /%2f|%5c|%25/i.test(url.pathname)) throw new Error("Unsupported product read URL");
    const paged = url.pathname === "/v1.0/me/people" || url.pathname === "/v1.0/groups"
        || url.pathname.endsWith("/children") || url.pathname.endsWith("/search(q='.loop')");
    const user = /^\/v1\.0\/users\/[^/]+$/.test(url.pathname), keys = new Set();
    for (const [key, value] of url.searchParams) {
        if (keys.has(key)) throw new Error("Duplicate product read query");
        keys.add(key);
        const allowed = key === "$top" && paged && /^\d+$/.test(value) && Number(value) >= 1 && Number(value) <= 100
            || key === "$select" && user && value === "id,displayName,mail,userPrincipalName,userType"
            || key === "$filter" && url.pathname === "/v1.0/groups" && value === "groupTypes/any(c:c eq 'Unified')";
        if (!allowed) throw new Error("Unsupported product read query");
    }
    let token = await acquireToken("https://graph.microsoft.com");
    if (token.status !== 200) {
        if (token.status === 422 || token.status === 425) throw new Error("Outlook authentication contract is unavailable");
        return {status: 401};
    }
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), 25000);
    try {
        const response = await fetch(url.href, {
            method: "GET", mode: "cors", credentials: "omit", redirect: "error",
            headers: {accept: "application/json", authorization: "Bearer " + token.accessToken},
            signal: controller.signal
        });
        token = null;
        return await readResponse(response);
    } finally {
        token = null;
        clearTimeout(timer);
        controller.abort();
    }
}
