async function(config, signal) {
    const base = new URL(config.base);
    if (location.origin !== base.origin || base.protocol !== "https:" || base.username || base.password
        || base.port || base.search || base.hash || !base.hostname.endsWith(".sharepoint.com")
        || !/^\/(?:sites|teams|personal)\/[^/]+$/.test(base.pathname)
        || /%2f|%5c|%25|\\|\.\./i.test(base.pathname)
        || !/^[0-9a-f]{8}(?:-[0-9a-f]{4}){3}-[0-9a-f]{12}$/.test(config.id))
        throw new Error("Unsupported SharePoint document target");
    const endpoint = base.href + "/_api/web/GetFileById(guid'" + config.id + "')";
    const request = async (url, init = {}) => fetch(url, {
        signal, credentials:"same-origin", redirect:"error", cache:"no-store", ...init
    });
    const metadata = async () => {
        const response = await request(endpoint + "?$select=UniqueId,Name,Length,ETag,TimeLastModified,LockedByUser/Id&$expand=LockedByUser", {
            headers:{accept:"application/json;odata=nometadata"}
        });
        if (!response.ok) return {status:response.status};
        const value = await response.json();
        const lock = value.LockedByUser;
        if (typeof value.Name !== "string" || !value.Name.toLowerCase().endsWith(".docx")
            || typeof value.ETag !== "string" || !value.ETag || value.ETag === "*" || /[\r\n]/.test(value.ETag)
            || typeof value.UniqueId !== "string" || value.UniqueId.toLowerCase() !== config.id
            || !/^\d+$/.test(String(value.Length)) || Number(value.Length) > 16000000
            || !(lock == null || lock?.["odata.null"] === true
                || (Number.isSafeInteger(lock?.Id) && lock.Id > 0)))
            throw new Error("Invalid Word document metadata");
        return {status:200,id:config.id,name:value.Name,size:Number(value.Length),
            etag:value.ETag,modified:value.TimeLastModified,locked:Number.isSafeInteger(lock?.Id) && lock.Id > 0};
    };
    if (config.direction === "metadata") return metadata();
    if (!Number.isSafeInteger(config.size) || config.size <= 0 || config.size > 16000000
        || typeof config.etag !== "string" || !config.etag || config.etag === "*" || /[\r\n]/.test(config.etag))
        throw new Error("Invalid document transfer bounds or version");
    if (config.direction === "download") {
        const response = await request(endpoint + "/$value", {headers:{"if-match":config.etag}});
        if (!response.ok || !response.body) throw new Error("Document download rejected");
        const declared = response.headers.get("content-length");
        if (declared !== null && (!/^\d+$/.test(declared) || Number(declared) !== config.size))
            throw new Error("Document size changed");
        const reader = response.body.getReader();
        let pending = null, offset = 0, position = 0;
        return {
            async read(maxBytes) {
                while (!pending || position === pending.length) {
                    const chunk = await reader.read();
                    if (chunk.done) {
                        if (offset !== config.size) throw new Error("Document download was incomplete");
                        return null;
                    }
                    if (!(chunk.value instanceof Uint8Array) || offset + chunk.value.length > config.size)
                        throw new Error("Document exceeded declared size");
                    pending = chunk.value;
                    position = 0;
                }
                if (maxBytes === 0) throw new Error("Document exceeded transfer ceiling");
                const bytes = pending.slice(position, position + maxBytes);
                position += bytes.length;
                offset += bytes.length;
                return bytes;
            },
            async close() { await reader.cancel(); pending = null; },
            async abort() { await reader.cancel(); pending = null; }
        };
    }
    if (config.direction !== "upload" || config.allow_mutation !== true
        || !/^__moff_doc_[a-z0-9_-]+$/i.test(config.result_key))
        throw new Error("Document writes require explicit mutation approval");
    let data = new Uint8Array(config.size), received = 0, committed = false;
    return {
        async write(bytes, offset) {
            if (committed || !(bytes instanceof Uint8Array) || bytes.length > 128000
                || offset !== received || received + bytes.length > data.length)
                throw new Error("Invalid document upload chunk");
            data.set(bytes, received);
            received += bytes.length;
        },
        async close() {
            if (committed || received !== config.size) throw new Error("Incomplete document upload");
            committed = true;
            window[config.result_key] = {stage:"digest",status:null};
            setTimeout(() => { delete window[config.result_key]; }, 120000);
            const context = await request(base.href + "/_api/contextinfo", {
                method:"POST", headers:{accept:"application/json;odata=nometadata"}
            });
            if (!context.ok) {
                window[config.result_key] = {stage:"digest",status:context.status};
                throw new Error("SharePoint request digest unavailable");
            }
            const value = await context.json();
            if (typeof value.FormDigestValue !== "string" || !value.FormDigestValue)
                throw new Error("Invalid SharePoint request digest");
            window[config.result_key] = {stage:"write",status:null};
            const response = await request(endpoint + "/$value", {
                method:"POST",headers:{"X-HTTP-Method":"PUT","X-RequestDigest":value.FormDigestValue,
                    "If-Match":config.etag,"Content-Type":"application/octet-stream"},
                body:data
            });
            window[config.result_key] = {stage:"write",status:response.status};
            data = null;
            if (!response.ok) throw new Error("Versioned document update rejected");
        },
        async abort() { data = null; }
    };
}
