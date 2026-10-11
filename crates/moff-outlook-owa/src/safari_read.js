async function(config, acquireToken, readResponse) {
    if (!["https://outlook.office.com", "https://outlook.cloud.microsoft"].includes(location.origin)) {
        throw new Error("Moff requires an HTTPS Outlook tab");
    }
    const allowed = ["FindItem", "FindConversation", "GetItem", "GetConversationItems",
        "GetCalendarFolders", "GetCalendarFoldersV2", "GetCalendarView"];
    const authOnly = config.action === "__MoffAuthStatus" && config.mutation !== true;
    if (authOnly && config.recovery_marker !== undefined
        && (!config.recovery_marker || sessionStorage.getItem("__moff_auth_recovery") !== config.recovery_marker)) {
        return {status: 409};
    }
    const writes = ["CreateItem", "UpdateItem", "UpdateCalendarEvent"];
    if (config.mutation === true) {
        if (config.allow_mutation !== true || !writes.includes(config.action)) throw new Error("Unsupported write action");
        const body = config.body?.Body;
        if (!body || body.__type !== config.action + "Request:#Exchange") throw new Error("Invalid write envelope");
        if (config.action === "CreateItem") {
            if (!Array.isArray(body.Items) || body.Items.length !== 1) throw new Error("Invalid item count");
            const item = body.Items[0];
            if (item.__type === "CalendarItem:#Exchange") {
                if (body.SendMeetingInvitations !== "SendToNone"
                    || (body.MessageDisposition !== undefined && body.MessageDisposition !== "SaveOnly")
                    || ["RequiredAttendees", "OptionalAttendees", "Resources"].some(k => item[k]?.length)) {
                    throw new Error("Invitations are prohibited");
                }
            } else if (!["Message:#Exchange", "ReplyToItem:#Exchange", "ReplyAllToItem:#Exchange"].includes(item.__type)
                || body.MessageDisposition !== "SaveOnly") throw new Error("Sending is prohibited");
        }
        if (config.action === "UpdateItem" && body.MessageDisposition !== "SaveOnly") throw new Error("Sending is prohibited");
        if (config.action === "UpdateCalendarEvent" && body.ShouldSendUpdateToAttendees !== false) throw new Error("Invitations are prohibited");
    } else if (!authOnly && !allowed.includes(config.action)) throw new Error("Unsupported read action");
    let token = await acquireToken(location.origin);
    if (token.status !== 200) {
        if (authOnly || token.status === 401) return {status: token.status};
        if (token.status === 425) throw new Error("Outlook is not loaded; sign in and wait for Mail");
        throw new Error("Outlook authentication contract changed; Moff needs an adapter update");
    }
    if (authOnly) {
        token = null;
        if (config.recovery_marker !== undefined) sessionStorage.removeItem("__moff_auth_recovery");
        return {status: 200};
    }
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), 25000);
    try {
        const response = await fetch("/owa/service.svc?action=" + config.action + "&app=Mail", {
            method: "POST", mode: "same-origin", credentials: "same-origin", redirect: "error",
            headers: {"content-type": "application/json", action: config.action,
                authorization: "Bearer " + token.accessToken},
            body: JSON.stringify(config.body), signal: controller.signal
        });
        token = null;
        return await readResponse(response);
    } finally {
        token = null;
        clearTimeout(timer);
        controller.abort();
    }
}
