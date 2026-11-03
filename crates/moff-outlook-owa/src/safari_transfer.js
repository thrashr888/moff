async function(config, acquireToken, readResponse, signal) {
    if (!["https://outlook.office.com", "https://outlook.cloud.microsoft", "https://outlook.live.com"].includes(location.origin))
        throw new Error("Transfers require an Outlook document");
    const url = new URL(config.url);
    const item = /^\/v1\.0\/me\/drive\/items\/[^/]+$/.test(url.pathname);
    const existing = /^\/v1\.0\/me\/drive\/items\/[^/]+\/content$/.test(url.pathname);
    const creating = /^\/v1\.0\/me\/drive\/(root|items\/[^/]+):\/[^/]+:\/content$/.test(url.pathname);
    if (url.origin !== "https://graph.microsoft.com" || url.username || url.password || url.hash
        || /%2f|%5c|%25/i.test(url.pathname)
        || url.pathname.split("/").some(part => /^(?:\.|%2e){1,2}$/i.test(part))
        || !Number.isSafeInteger(config.size) || config.size < 0 || config.size > 250000000)
        throw new Error("Invalid transfer configuration");
    const uploading = config.direction === "upload";
    if (uploading ? (!existing && !creating || config.allow_mutation !== true
        || existing && (typeof config.etag !== "string" || !config.etag || /[\r\n]/.test(config.etag))
        || creating && url.search !== "?@microsoft.graph.conflictBehavior=fail"
        || existing && url.search || !/^__moff_upload_[a-z0-9_-]+$/i.test(config.result_key))
        : (config.direction !== "download" || !item || url.search))
        throw new Error("Invalid transfer route or mutation gate");
    let stage = "authentication", status = null;
    const fail = error => {
        if (uploading) {
            window[config.result_key + "_failure"] = {stage,status};
            setTimeout(() => { delete window[config.result_key + "_failure"]; }, 120000);
        }
        throw error;
    };
    const guarded = operation => async (...args) => {
        try { return await operation(...args); } catch (error) { return fail(error); }
    };
    let token = await guarded(acquireToken)("https://graph.microsoft.com");
    if (token.status !== 200) return fail(new Error("Product authentication unavailable"));
    const headers = {authorization: "Bearer " + token.accessToken, accept: "application/json",
        ...(config.etag ? {"if-match": config.etag} : {})};
    token = null;
    const request = guarded(async (target, init) => {
        status = null;
        const response = await fetch(target, {mode:"cors", credentials:"omit", redirect:"error", signal, ...init});
        status = response.status;
        if (!response.ok) throw new Error("File request rejected");
        return response;
    });
    const jsonResponse = guarded(async response => {
        const result = await readResponse(response);
        if (result.status !== response.status || typeof result.body !== "string") throw new Error("Invalid file response");
        const value = JSON.parse(result.body);
        if (value.error) throw new Error("File request rejected");
        return value;
    });
    const microsoftUrl = value => {
        const target = new URL(value);
        if (target.protocol !== "https:" || target.username || target.password || target.hash
            || !(target.hostname.endsWith(".sharepoint.com") || target.hostname.endsWith(".1drv.com")
                || target.hostname.endsWith(".onedrive.com") || target.hostname === "onedrive.live.com"))
            return fail(new Error("Unsupported Microsoft transfer location"));
        return target.href;
    };
    if (!uploading) {
        const metadata = await jsonResponse(await request(url.href, {headers}));
        if (!metadata.file || metadata.size !== config.size) throw new Error("File size changed");
        const downloadUrl = microsoftUrl(metadata["@microsoft.graph.downloadUrl"]);
        const rangeBytes = 3276800;
        let reader = null, pending = null, offset = 0, position = 0, rangeEnd = 0;
        const openRange = async () => {
            rangeEnd = Math.min(offset + rangeBytes, config.size);
            const response = await request(downloadUrl, {method:"GET",
                ...(config.size ? {headers:{Range:`bytes=${offset}-${rangeEnd - 1}`}} : {})});
            // SharePoint's CORS response may not expose Content-Range; length is checked independently.
            const contentRange = response.headers.get("content-range"), length = response.headers.get("content-length");
            if (!response.body || (response.status !== 206 && !(response.status === 200 && config.size <= rangeBytes && offset === 0))
                || response.status === 206 && contentRange !== null && contentRange !== `bytes ${offset}-${rangeEnd - 1}/${config.size}`
                || length !== null && (!/^\d+$/.test(length) || Number(length) !== rangeEnd - offset))
                throw new Error("Invalid bounded download range");
            reader = response.body.getReader();
        };
        await openRange();
        return {
            async read(maxBytes) {
                while (!pending || position === pending.length) {
                    const chunk = await reader.read();
                    if (chunk.done) {
                        if (offset !== rangeEnd) throw new Error("Download range length mismatch");
                        await reader.cancel();
                        if (offset === config.size) return null;
                        await openRange();
                        continue;
                    }
                    if (!(chunk.value instanceof Uint8Array) || chunk.value.length > rangeBytes
                        || offset + chunk.value.length > rangeEnd) throw new Error("Download buffer or length exceeded");
                    pending = chunk.value;
                    position = 0;
                }
                if (maxBytes === 0) throw new Error("Download exceeds ceiling");
                const bytes = pending.slice(position, position + maxBytes);
                position += bytes.length;
                offset += bytes.length;
                return bytes;
            },
            async close() { await reader.cancel(); pending = null; },
            async abort() { await reader.cancel(); pending = null; }
        };
    }
    let uploadUrl, result;
    const serviceBytes = 3276800; // Ten Graph 320 KiB ranges; independent of Cider's transport chunks.
    let buffer = new Uint8Array(Math.min(serviceBytes, config.size)), used = 0, received = 0, sent = 0;
    const verifyResult = value => {
        if (typeof value.id !== "string" || !value.id || value.size !== config.size)
            throw new Error("Upload acknowledgement could not be verified");
        return {id:value.id, size:value.size, name:value.name, eTag:value.eTag, verified:false, acknowledged:true};
    };
    if (config.size === 0) {
        stage = "empty upload";
        result = verifyResult(await jsonResponse(await request(url.href, {
            method:"PUT", headers:{...headers, "content-type":"application/octet-stream"}, body:new Uint8Array()
        })));
    } else {
        stage = "upload session creation";
        const sessionUrl = new URL(url.href);
        sessionUrl.pathname = sessionUrl.pathname.replace(/\/content$/, "/createUploadSession");
        sessionUrl.search = "";
        const session = await jsonResponse(await request(sessionUrl.href, {
            method:"POST", headers:{...headers, "content-type":"application/json"},
            body:JSON.stringify({item:{"@microsoft.graph.conflictBehavior":creating ? "fail" : "replace"}})
        }));
        uploadUrl = microsoftUrl(session.uploadUrl);
    }
    return {
        write: guarded(async (bytes, offset) => {
            stage = "upload range";
            if (!(bytes instanceof Uint8Array) || bytes.length > 128000 || offset !== received
                || received + bytes.length > config.size) throw new Error("Invalid upload chunk");
            let position = 0;
            while (position < bytes.length) {
                const count = Math.min(buffer.length - used, bytes.length - position);
                buffer.set(bytes.subarray(position, position + count), used);
                used += count;
                position += count;
                received += count;
                if (used === buffer.length || received === config.size) {
                    const end = sent + used;
                    const response = await request(uploadUrl, {method:"PUT",
                        headers:{"content-type":"application/octet-stream",
                            "content-range":`bytes ${sent}-${end - 1}/${config.size}`},
                        body:buffer.subarray(0, used)});
                    const value = await jsonResponse(response);
                    if (end === config.size) {
                        if (![200,201].includes(response.status)) throw new Error("Missing final upload acknowledgement");
                        result = verifyResult(value);
                    } else if (response.status !== 202 || !Array.isArray(value.nextExpectedRanges)
                        || value.nextExpectedRanges.length !== 1
                        || !new RegExp(`^${end}-(?:${config.size - 1})?$`).test(value.nextExpectedRanges[0])) {
                        throw new Error("Unknown upload range outcome");
                    }
                    sent = end;
                    used = 0;
                }
            }
        }),
        close: guarded(async () => {
            stage = "upload finalization";
            if (received !== config.size || !result) throw new Error("Incomplete upload");
            window[config.result_key] = result;
            setTimeout(() => { delete window[config.result_key]; }, 120000);
            buffer = null;
            uploadUrl = null;
        }),
        abort() {
            // Do not retry or issue another write after an ambiguous range request.
            buffer = null;
            uploadUrl = null;
            delete window[config.result_key];
        }
    };
}
