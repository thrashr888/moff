async function(config, acquireToken, readResponse) {
    if (!["https://outlook.office.com", "https://outlook.cloud.microsoft", "https://outlook.live.com"].includes(location.origin)) {
        throw new Error("Moff requires an HTTPS Outlook tab");
    }
    const allowed = ["FindFolder", "GetFolder", "FindItem", "FindConversation", "GetItem", "GetConversationItems",
        "GetCalendarFolders", "GetCalendarFoldersV2", "GetCalendarView"];
    const authOnly = config.action === "__MoffAuthStatus" && config.mutation !== true;
    if (authOnly && config.recovery_marker !== undefined
        && (!config.recovery_marker || sessionStorage.getItem("__moff_auth_recovery") !== config.recovery_marker)) {
        return {status: 409};
    }
    const writes = ["CreateItem", "UpdateItem", "UpdateCalendarEvent", "DeleteItem", "MoveItem"];
    if (config.mutation === true) {
        if (config.allow_mutation !== true || !writes.includes(config.action)) throw new Error("Unsupported write action");
        const body = config.body?.Body;
        if (!body || body.__type !== config.action + "Request:#Exchange") throw new Error("Invalid write envelope");
        if (config.action === "MoveItem") {
            const folder = body.ToFolderId?.BaseFolderId;
            if (body.ReturnNewItemIds !== true || folder?.__type !== "FolderId:#Exchange"
                || typeof folder.Id !== "string" || !folder.Id.trim()
                || !Array.isArray(body.ItemIds) || body.ItemIds.length !== 1
                || !body.ItemIds[0].Id || !body.ItemIds[0].ChangeKey) {
                throw new Error("Only a single versioned move to an explicit folder is allowed");
            }
        }
        if (config.action === "CreateItem") {
            if (!Array.isArray(body.Items) || body.Items.length !== 1) throw new Error("Invalid item count");
            const item = body.Items[0];
            if (item.__type === "CalendarItem:#Exchange") {
                if (body.SendMeetingInvitations !== "SendToNone"
                    || (body.MessageDisposition !== undefined && body.MessageDisposition !== "SaveOnly")
                    || ["RequiredAttendees", "OptionalAttendees", "Resources"].some(k => item[k]?.length)) {
                    throw new Error("Invitations are prohibited");
                }
            } else if (["AcceptItem:#Exchange", "TentativelyAcceptItem:#Exchange", "DeclineItem:#Exchange"].includes(item.__type)) {
                const allowedKeys = ["__type", "ReferenceItemId", "Body", "ProposedStart", "ProposedEnd"];
                const allowedBodyKeys = ["__type", "MessageDisposition", "ItemShape", "Items"];
                const hasProposal = item.ProposedStart !== undefined || item.ProposedEnd !== undefined;
                const proposalStart = Date.parse(item.ProposedStart);
                const proposalEnd = Date.parse(item.ProposedEnd);
                if (body.MessageDisposition !== "SendAndSaveCopy" || body.Items.length !== 1
                    || Object.keys(body).some(k => !allowedBodyKeys.includes(k))
                    || Object.keys(item).some(k => !allowedKeys.includes(k))
                    || !item.ReferenceItemId?.Id || !item.ReferenceItemId?.ChangeKey
                    || Object.keys(item.ReferenceItemId).some(k => !["__type", "Id", "ChangeKey"].includes(k))
                    || (item.Body !== undefined
                        && (item.Body.__type !== "BodyContentType:#Exchange"
                            || item.Body.BodyType !== "Text" || typeof item.Body.Value !== "string"))
                    || (hasProposal && (!["DeclineItem:#Exchange", "TentativelyAcceptItem:#Exchange"].includes(item.__type)
                        || !item.ProposedStart || !item.ProposedEnd || !Number.isFinite(proposalStart)
                        || !Number.isFinite(proposalEnd) || proposalEnd <= proposalStart))) {
                    throw new Error("Invalid meeting response");
                }
            } else if (!["Message:#Exchange", "ReplyToItem:#Exchange", "ReplyAllToItem:#Exchange","Contact:#Exchange"].includes(item.__type)
                || body.MessageDisposition !== "SaveOnly") throw new Error("Sending is prohibited");
            if (item.__type === "Contact:#Exchange" && body.SavedItemFolderId?.BaseFolderId?.Id !== "contacts")
                throw new Error("Contacts must be created in the personal contacts folder");
        }
        if (config.action === "UpdateItem" && body.MessageDisposition !== "SaveOnly") throw new Error("Sending is prohibited");
        if (config.action === "UpdateCalendarEvent") {
            if (body.ShouldSendUpdateToAttendees === true) {
                const itemId = body.ItemChange?.ItemId;
                const eventId = body.EventId;
                const updates = body.ItemChange?.Updates;
                const fields = Array.isArray(updates)
                    ? updates.map(update => update.Path?.FieldURI).sort()
                    : [];
                const start = updates?.find(update => update.Path?.FieldURI === "Start")?.Item?.Start;
                const end = updates?.find(update => update.Path?.FieldURI === "End")?.Item?.End;
                if (body.EventScope !== 0 || body.TargetAudience !== 0 || fields.join(",") !== "End,Start"
                    || !itemId?.Id || !itemId?.ChangeKey || eventId?.Id !== itemId.Id
                    || eventId?.ChangeKey !== itemId.ChangeKey || !start || !end
                    || !Number.isFinite(Date.parse(start)) || !Number.isFinite(Date.parse(end))
                    || Date.parse(end) <= Date.parse(start)) {
                    throw new Error("Only an explicit, versioned meeting-time move may notify attendees");
                }
            } else if (body.ShouldSendUpdateToAttendees !== false) {
                throw new Error("Invitations are prohibited");
            }
        }
        if (config.action === "DeleteItem" && (body.DeleteType !== "MoveToDeletedItems"
            || !["SendToNone", "SendToAllAndSaveCopy"].includes(body.SendMeetingCancellations)
            || body.SuppressReadReceipts !== true
            || !Array.isArray(body.ItemIds) || body.ItemIds.length !== 1
            || !body.ItemIds[0].Id || !body.ItemIds[0].ChangeKey)) throw new Error("Only a single versioned recycle or attendee-notifying cancellation is allowed");
    } else if (!authOnly && !allowed.includes(config.action)) throw new Error("Unsupported read action");
    const controller = new AbortController();
    let token = null;
    let expired = false, dispatched = false, timer;
    const deadline = new Promise(resolve => {
        timer = setTimeout(() => {
            expired = true;
            controller.abort();
            resolve({status: 0, failure: "timeout", dispatched});
        }, config.timeout_ms ?? 25000);
    });
    try {
        return await Promise.race([deadline, (async () => {
            token = await acquireToken(location.origin);
            if (expired) { token = null; return {status: 0, failure: "timeout", dispatched}; }
            if (token.status !== 200) return authOnly
                ? {status: token.status} : {status: token.status, stage: "authentication"};
            if (authOnly) {
                token = null;
                if (config.recovery_marker !== undefined) sessionStorage.removeItem("__moff_auth_recovery");
                return {status: 200};
            }
            dispatched = true;
            const response = await fetch("/owa/service.svc?action=" + config.action + "&app=Mail", {
                method: "POST", mode: "same-origin", credentials: "same-origin", redirect: "error",
                headers: {"content-type": "application/json", action: config.action,
                    authorization: token.authorization},
                body: JSON.stringify(config.body), signal: controller.signal
            });
            token = null;
            return await readResponse(response);
        })()]);
    } catch (error) {
        if (expired || error?.name === "AbortError") return {status: 0, failure: "timeout", dispatched};
        if (error?.code === "moff_response_too_large") return {status: 0, failure: "response_too_large", dispatched};
        if (error instanceof TypeError || error?.name === "TypeError") return {status: 0, failure: "network", dispatched};
        throw error;
    } finally {
        token = null;
        clearTimeout(timer);
        controller.abort();
    }
}
