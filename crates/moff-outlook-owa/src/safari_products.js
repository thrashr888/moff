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
    const method = config.method || "GET";
    const paths = [
        /^\/v1\.0\/me\/(people|manager|directReports)$/,
        /^\/v1\.0\/me$/,
        /^\/v1\.0\/me\/contacts(\/[^/]+)?$/,
        /^\/v1\.0\/users(\/[^/]+(\/(manager|directReports))?)?$/,
        /^\/v1\.0\/groups(\/[^/]+(\/(owners|members))?)?$/,
        /^\/v1\.0\/groups\/[^/]+\/members\/(\$ref|[^/]+\/\$ref)$/,
        /^\/v1\.0\/me\/drive\/(root|items\/[^/]+)\/children$/,
        /^\/v1\.0\/me\/drive\/items\/[^/]+$/,
        /^\/v1\.0\/(?:me\/drive|drives\/[^/]+)\/items\/[^/]+\/(?:invite|permissions(?:\/[^/]+)?)$/,
        /^\/v1\.0\/drives\/[^/]+\/items\/[^/]+$/,
        /^\/v1\.0\/me\/drive\/root\/search\(q='[^/]*'\)$/,
        /^\/v1\.0\/drives\/[^/]+\/root\/search\(q='[^/]*'\)$/,
        /^\/v1\.0\/sites\/[^/]+(?:\/(?:sites|teams)\/[^/]+|\/drives)?$/,
    ];
    if (url.origin !== "https://graph.microsoft.com" || url.username || url.password || url.hash
        || !paths.some(pattern => pattern.test(url.pathname))
        || /%2f|%5c|%25/i.test(url.pathname)) throw new Error("Unsupported product read URL");
    const paged = url.pathname === "/v1.0/me/people" || url.pathname === "/v1.0/me/contacts" || url.pathname === "/v1.0/groups" || url.pathname === "/v1.0/users"
        || /^\/v1\.0\/groups\/[^/]+\/(owners|members)$/.test(url.pathname)
        || /^\/v1\.0\/(?:me\/drive|drives\/[^/]+)\/items\/[^/]+\/permissions$/.test(url.pathname)
        || /^\/v1\.0\/sites\/[^/]+\/drives$/.test(url.pathname)
        || /^\/v1\.0\/drives\/[^/]+\/root\/search\(q='[^/]*'\)$/.test(url.pathname)
        || url.pathname.endsWith("/children") || /^\/v1\.0\/me\/drive\/root\/search\(q='[^/]*'\)$/.test(url.pathname);
    const user = /^\/v1\.0\/users\/[^/]+$/.test(url.pathname), keys = new Set();
    for (const [key, value] of url.searchParams) {
        if (keys.has(key)) throw new Error("Duplicate product read query");
        keys.add(key);
        const allowed = key === "$top" && paged && /^\d+$/.test(value) && Number(value) >= 1 && Number(value) <= 100
            || key === "$select" && user && value === "id,displayName,mail,userPrincipalName,userType"
            || key === "$filter" && url.pathname === "/v1.0/groups" && value === "groupTypes/any(c:c eq 'Unified')"
            || key === "$skiptoken" && paged && value.length > 0 && value.length <= 4096
            || key === "$skip" && paged && /^\d+$/.test(value) && Number(value) <= 4294967295;
        if (!allowed) throw new Error("Unsupported product read query");
    }
    const write = method !== "GET";
    const objectFields = (value, fields) => {
        if (!value || Array.isArray(value) || typeof value !== "object" || !Object.keys(value).length
            || Object.keys(value).some(key => !fields.includes(key))) throw new Error("Invalid product write body");
    };
    const strings = (value, fields) => {
        objectFields(value, fields);
        if (Object.values(value).some(item => typeof item !== "string")) throw new Error("Invalid string fields");
    };
    const segment = value => typeof value === "string" && value.trim()
        && !/[\\/%\x00-\x1f]/.test(value) && value !== "." && value !== "..";
    if (write) {
        if (config.allow_mutation !== true || !["POST", "PATCH", "DELETE"].includes(method)
            || url.search
            || JSON.stringify(config.body || {}).length > 200000) throw new Error("Product writes require an explicit gate");
        const contact = /^\/v1\.0\/me\/contacts\/[^/]+$/.test(url.pathname);
        const drive = /^\/v1\.0\/me\/drive\/items\/[^/]+$/.test(url.pathname);
        const group = /^\/v1\.0\/groups\/[^/]+$/.test(url.pathname);
        const folder = /^\/v1\.0\/me\/drive\/(root|items\/[^/]+)\/children$/.test(url.pathname);
        const memberAdd = /^\/v1\.0\/groups\/[^/]+\/members\/\$ref$/.test(url.pathname);
        const memberRemove = /^\/v1\.0\/groups\/[^/]+\/members\/[^/]+\/\$ref$/.test(url.pathname);
        const invite = /^\/v1\.0\/(?:me\/drive|drives\/[^/]+)\/items\/[^/]+\/invite$/.test(url.pathname);
        if (method === "DELETE") {
            if (!(contact || drive || group || memberRemove) || config.body != null) throw new Error("Unsupported product delete");
        } else if (method === "POST" && invite) {
            objectFields(config.body, ["recipients","roles","requireSignIn","sendInvitation","retainInheritedPermissions"]);
            const body = config.body;
            if (!Array.isArray(body.recipients) || body.recipients.length !== 1
                || JSON.stringify(body.roles) !== '["read"]' || body.requireSignIn !== true
                || typeof body.sendInvitation !== "boolean" || body.retainInheritedPermissions !== true)
                throw new Error("Sharing requires one named recipient with signed-in read access");
            strings(body.recipients[0], ["email"]);
            const email = body.recipients[0].email;
            if (typeof email !== "string" || email.length > 254 || /[\s\x00-\x1f\x7f]/.test(email)
                || !/^[a-z0-9.!#$%&'*+\-/=?^_`{|}~]+@[a-z0-9](?:[a-z0-9-]*[a-z0-9])?(?:\.[a-z0-9](?:[a-z0-9-]*[a-z0-9])?)*$/i.test(email))
                throw new Error("Sharing requires one explicit email address");
        } else if (method === "POST" && memberAdd) {
            strings(config.body, ["@odata.id"]);
            const member = new URL(config.body["@odata.id"]);
            if (member.origin !== "https://graph.microsoft.com" || member.username || member.password || member.search || member.hash
                || !/^\/v1\.0\/users\/[^/]+$/.test(member.pathname) || /%2f|%5c|%25/i.test(member.pathname)) throw new Error("Only directory users can be group members");
        } else if ((method === "POST" && url.pathname === "/v1.0/me/contacts") || (method === "PATCH" && contact)) {
            objectFields(config.body, ["givenName","surname","displayName","companyName","jobTitle","department","officeLocation","mobilePhone","businessPhones","homePhones","emailAddresses","personalNotes"]);
            for (const [key, value] of Object.entries(config.body)) {
                if (["businessPhones","homePhones"].includes(key)) {
                    if (!Array.isArray(value) || value.some(item => typeof item !== "string")) throw new Error("Invalid phones");
                } else if (key === "emailAddresses") {
                    if (!Array.isArray(value)) throw new Error("Invalid addresses");
                    for (const address of value) {
                        strings(address, ["name","address"]);
                        if (!address.address?.includes("@") || /[\x00-\x1f]/.test(address.address)) throw new Error("Invalid address");
                    }
                } else if (typeof value !== "string") throw new Error("Invalid contact field");
            }
        } else if (method === "POST" && folder) {
            objectFields(config.body, ["name","folder","@microsoft.graph.conflictBehavior"]);
            if (!segment(config.body.name) || !config.body.folder || Array.isArray(config.body.folder)
                || typeof config.body.folder !== "object" || Object.keys(config.body.folder).length
                || config.body["@microsoft.graph.conflictBehavior"] !== "fail") throw new Error("Invalid folder body");
        } else if (method === "PATCH" && drive) {
            objectFields(config.body, ["name","description","parentReference"]);
            for (const [key,value] of Object.entries(config.body)) {
                if (key === "parentReference") {
                    strings(value, ["id"]);
                    if (!segment(value.id)) throw new Error("Invalid parent folder ID");
                } else if (typeof value !== "string") throw new Error("Invalid drive field");
            }
            if (config.body.name !== undefined && !segment(config.body.name)) throw new Error("Invalid drive name");
        } else if ((method === "POST" && url.pathname === "/v1.0/groups") || (method === "PATCH" && group)) {
            objectFields(config.body, method === "POST"
                ? ["displayName","description","mailNickname","visibility","groupTypes","mailEnabled","securityEnabled"]
                : ["displayName","description","visibility"]);
            for (const key of ["displayName","description","mailNickname","visibility"]) {
                if (config.body[key] !== undefined && typeof config.body[key] !== "string") throw new Error("Invalid group field");
            }
            if (config.body.displayName !== undefined && (!config.body.displayName.trim() || config.body.displayName.length > 256))
                throw new Error("Invalid group name");
            if (config.body.visibility !== undefined && !["Private","Public"].includes(config.body.visibility)) throw new Error("Invalid visibility");
            if (method === "POST" && (!config.body.displayName || !/^[a-z0-9_-]{1,64}$/i.test(config.body.mailNickname)
                || JSON.stringify(config.body.groupTypes) !== '["Unified"]'
                || config.body.mailEnabled !== true || config.body.securityEnabled !== false)) throw new Error("Only Microsoft 365 groups can be created");
        } else throw new Error("Unsupported product write route");
        if (drive && !config.etag) throw new Error("Drive writes require the current entity tag");
    } else if (config.body != null || config.etag != null || config.bytes != null || url.pathname.endsWith("/invite")
        || url.pathname.endsWith("/$ref")) throw new Error("Read cannot carry write fields");
    if (config.bytes != null || config.download === true) throw new Error("Use the streamed transfer adapter");
    if (config.etag != null && (typeof config.etag !== "string" || !config.etag || /[\r\n]/.test(config.etag))) throw new Error("Invalid entity tag");
    let token = await acquireToken("https://graph.microsoft.com");
    if (token.status !== 200) {
        if (token.status === 422 || token.status === 425) throw new Error("Outlook authentication contract is unavailable");
        return {status: 401};
    }
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), 25000);
    try {
        const response = await fetch(url.href, {
            method, mode: "cors", credentials: "omit", redirect: "error",
            headers: {accept: "application/json", authorization: "Bearer " + token.accessToken,
                ...(config.body != null ? {"content-type":"application/json"} : {}),
                ...(config.etag ? {"if-match":config.etag} : {})},
            ...(config.body != null ? {body:JSON.stringify(config.body)} : {}),
            signal: controller.signal
        });
        token = null;
        if (response.status === 204) return {status:204, body:""};
        const result = await readResponse(response);
        if (result.body !== undefined) {
            const value = JSON.parse(result.body);
            const scrub = item => {
                if (!item || typeof item !== "object") return;
                for (const key of Object.keys(item)) {
                    if (key.startsWith("@microsoft.graph.downloadUrl")) delete item[key];
                }
                for (const child of Object.values(item)) scrub(child);
            };
            scrub(value);
            result.body = JSON.stringify(value);
        }
        return result;
    } finally {
        token = null;
        clearTimeout(timer);
        controller.abort();
    }
}
