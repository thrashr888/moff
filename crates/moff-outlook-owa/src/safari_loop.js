async function(config, acquireToken, readResponse) {
    if (!["https://outlook.office.com", "https://outlook.cloud.microsoft", "https://outlook.live.com"].includes(location.origin))
        throw new Error("Loop requires an Outlook document");
    const base = "https://substrate.office.com/recommended/api/v1.1/loop";
    const validatePod = id => {
        if (typeof id !== "string" || id.length > 4096 || !/^[A-Za-z0-9+/]+={0,2}$/.test(id))
            throw new Error("Workspace lookup requires a PodId");
        const decoded = atob(id), parts = decoded.split("|");
        if (btoa(decoded) !== id || parts.length !== 4 || parts[0] !== "ODSP"
            || !/^[a-z0-9.-]+$/.test(parts[1]) || !parts[2].startsWith("b!") || !parts[3])
            throw new Error("Invalid workspace PodId");
        return id;
    };
    if (!["list","get","create","delete"].includes(config.action)) throw new Error("Unsupported workspace action");
    const mutation = ["create","delete"].includes(config.action);
    if (mutation && config.allow_mutation !== true) throw new Error("Workspace writes require a gate");
    if (["get","delete"].includes(config.action)) validatePod(config.id);
    if (config.action === "list" && (!Number.isInteger(config.limit) || config.limit < 1 || config.limit > 100
        || config.query != null && (typeof config.query !== "string" || !config.query.trim() || config.query.length > 4096)))
        throw new Error("Invalid workspace list bounds");
    if (config.mru_id != null && !/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(config.mru_id))
        throw new Error("Invalid workspace MRU ID");
    let token = await acquireToken("https://substrate.office.com");
    if (token.status !== 200) return {status:token.status};
    const claims = JSON.parse(atob(token.accessToken.split(".")[1].replace(/-/g,"+").replace(/_/g,"/")));
    const guid = value => typeof value === "string" && /^[0-9a-f-]{36}$/i.test(value);
    if (!guid(claims.oid) || !guid(claims.tid)) throw new Error("Missing mailbox anchor claims");
    const headers = {authorization:"Bearer " + token.accessToken,
        "X-AnchorMailbox":`Oid:${claims.oid}@${claims.tid}`, accept:"application/json"};
    token = null;
    const controller = new AbortController(), timer = setTimeout(() => controller.abort(), 55000);
    const request = async (url, method = "GET", body, authHeaders = headers) => {
        const response = await fetch(url, {method, mode:"cors", credentials:"omit", redirect:"error",
            headers:{...authHeaders, "Client-Request-Id":crypto.randomUUID(),
                ...(body !== undefined ? {"Content-Type":"application/json"} : {})},
            ...(body !== undefined ? {body:JSON.stringify(body)} : {}), signal:controller.signal});
        const result = await readResponse(response);
        return {response, result, value:result.body ? JSON.parse(result.body) : null};
    };
    const normalize = workspace => {
        const id = validatePod(workspace.mfs_info?.pod_id);
        return {...workspace, source_id:workspace.id, id};
    };
    const success = value => ({status:200,body:JSON.stringify(value)});
    try {
        if (config.action === "list") {
            const {response, result, value} = await request(base + "/recent");
            if (!response.ok) return result;
            if (!Array.isArray(value.workspaces) || typeof value.is_complete !== "boolean") throw new Error("Invalid workspace list");
            let workspaces = value.workspaces.map(normalize);
            if (config.query != null) {
                const query = config.query.toLowerCase();
                workspaces = workspaces.filter(w => typeof w.title === "string" && w.title.toLowerCase().includes(query));
            }
            const complete = value.is_complete && workspaces.length <= config.limit;
            return success({workspaces:workspaces.slice(0,config.limit), is_complete:complete,
                scope:"recent workspaces only", scanned:value.workspaces.length});
        }
        const workspaceUrl = base + "/workspaces/" + encodeURIComponent(config.id);
        if (config.action === "get") {
            const {response, result, value} = await request(workspaceUrl);
            if (!response.ok) return result;
            if (!Array.isArray(value.workspaces) || value.workspaces.length !== 1) throw new Error("Invalid workspace response");
            return success({workspaces:value.workspaces.map(normalize)});
        }
        if (config.action === "delete") {
            const before = await request(workspaceUrl);
            if (!before.response.ok) return before.result;
            const workspaces = before.value?.workspaces;
            if (!Array.isArray(workspaces) || workspaces.length !== 1
                || normalize(workspaces[0]).id !== config.id) throw new Error("Workspace identity not established");
            const identities = [claims.upn, claims.preferred_username, claims.unique_name]
                .filter(v => typeof v === "string" && v).map(v => v.toLowerCase());
            if (!Array.isArray(workspaces[0].permissions) || !workspaces[0].permissions.some(p =>
                typeof p.role === "string" && p.role.toLowerCase() === "owner"
                && typeof p.upn === "string" && identities.includes(p.upn.toLowerCase())))
                return {status:403,body:JSON.stringify({error:"Workspace deletion requires current-user Owner permission"})};
            const deleted = await request(workspaceUrl, "DELETE", config.mru_id ? {mru_id:config.mru_id} : {});
            if (!deleted.response.ok) return deleted.result;
            if (deleted.value?.response_code !== 10010) throw new Error("Unknown workspace delete outcome");
            const after = await request(workspaceUrl);
            if (after.response.status !== 404) throw new Error("Workspace deletion not verified");
            return success({id:config.id, permanent:true, acknowledged:true, verified:true});
        }
        const created = await request(base + "/workspaces", "POST", {type:"tenantOwned"});
        if (!created.response.ok) return created.result;
        const workspace = created.value?.workspace, driveId = workspace?.id;
        if (created.response.status !== 201 || typeof driveId !== "string" || !/^b![A-Za-z0-9_-]+$/.test(driveId))
            throw new Error("Unknown workspace creation outcome");
        // Creation has already happened. Preserve its drive ID if PodId derivation fails.
        try {
            const site = new URL(workspace.sharepoint_info.site_url);
            if (site.protocol !== "https:" || !site.hostname.endsWith(".sharepoint.com") || site.username || site.password)
                throw new Error("Invalid created workspace site");
            const graphToken = await acquireToken("https://graph.microsoft.com");
            if (graphToken.status !== 200) throw new Error("Graph authentication unavailable");
            const root = await request("https://graph.microsoft.com/v1.0/drives/" + encodeURIComponent(driveId) + "/root?$select=id",
                "GET", undefined, {authorization:"Bearer " + graphToken.accessToken});
            if (!root.response.ok || typeof root.value?.id !== "string" || !root.value.id || root.value.id.includes("|"))
                throw new Error("Missing created workspace root");
            const id = validatePod(btoa(`ODSP|${site.hostname.toLowerCase()}|${driveId}|${root.value.id}`));
            return success({id, drive_id:driveId, sharepoint_info:workspace.sharepoint_info,
                acknowledged:true, verified:true, title_supported:false});
        } catch {
            return success({drive_id:driveId, verification_failed:true});
        }
    } finally {
        clearTimeout(timer);
        controller.abort();
    }
}
