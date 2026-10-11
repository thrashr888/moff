async function(config) {
    if (!["https://outlook.office.com", "https://outlook.cloud.microsoft"].includes(location.origin)) {
        throw new Error("Moff requires an HTTPS Outlook tab");
    }
    const allowed = ["FindItem", "FindConversation", "GetItem", "GetConversationItems",
        "GetCalendarFolders", "GetCalendarFoldersV2", "GetCalendarView"];
    if (!allowed.includes(config.action)) throw new Error("Unsupported read action");
    if (!window.webpackChunkOwa) throw new Error("Outlook is not loaded; sign in and wait for Mail");
    let runtime;
    const chunks = window.webpackChunkOwa;
    const marker = "__moff_read_" + crypto.randomUUID();
    chunks.push([[marker], {}, value => { runtime = value; }]);
    chunks.pop();
    if (!runtime || !runtime.m[121455] || !runtime.m[429896]) {
        throw new Error("Outlook authentication contract changed; Moff needs an adapter update");
    }
    const auth = runtime(121455);
    const mailbox = runtime(429896);
    if (typeof auth.ij !== "function" || typeof mailbox.A !== "function") {
        throw new Error("Outlook authentication contract changed; Moff needs an adapter update");
    }
    let token = await auth.ij("MoffRead", mailbox.A(), {resource: location.origin});
    if (!token || !token.AccessToken || token.TokenResultCode !== 0) {
        return {status: 401};
    }
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), 25000);
    try {
        const response = await fetch("/owa/service.svc?action=" + config.action + "&app=Mail", {
            method: "POST", mode: "same-origin", credentials: "same-origin", redirect: "error",
            headers: {"content-type": "application/json", action: config.action,
                authorization: "Bearer " + token.AccessToken},
            body: JSON.stringify(config.body), signal: controller.signal
        });
        token = null;
        if (!response.ok) return {status: response.status};
        const reader = response.body.getReader();
        const decoder = new TextDecoder();
        let text = "", bytes = 0;
        try {
            while (true) {
                const part = await reader.read();
                if (part.done) break;
                bytes += part.value.length;
                if (bytes > 900000) throw new Error("Outlook response exceeds Moff's browser limit");
                text += decoder.decode(part.value, {stream: true});
            }
            text += decoder.decode();
        } finally { await reader.cancel(); }
        return {status: response.status, body: text};
    } finally {
        token = null;
        clearTimeout(timer);
        controller.abort();
    }
}
