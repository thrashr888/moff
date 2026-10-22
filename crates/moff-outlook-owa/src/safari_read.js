async function(config, acquireToken, readResponse) {
    if (!["https://outlook.office.com", "https://outlook.cloud.microsoft", "https://outlook.live.com"].includes(location.origin)) {
        throw new Error("Moff requires an HTTPS Outlook tab");
    }
    const allowed = ["FindFolder", "GetFolder", "FindItem", "FindConversation", "GetItem", "GetAttachment", "GetConversationItems",
        "GetCalendarFolders", "GetCalendarFoldersV2", "GetCalendarView"];
    const authOnly = config.action === "__MoffAuthStatus" && config.mutation !== true;
    if (authOnly && config.recovery_marker !== undefined
        && (!config.recovery_marker || sessionStorage.getItem("__moff_auth_recovery") !== config.recovery_marker)) {
        return {status: 409};
    }
    const writes = ["CreateItem", "UpdateItem", "UpdateCalendarEvent", "DeleteItem", "MoveItem"];
    const onlyKeys = (value, keys) => value && typeof value === "object"
        && !Array.isArray(value) && Object.keys(value).every(key => keys.includes(key));
    if (config.mutation === true) {
        if (config.allow_mutation !== true || !writes.includes(config.action)) throw new Error("Unsupported write action");
        const body = config.body?.Body;
        if (!body || body.__type !== config.action + "Request:#Exchange") throw new Error("Invalid write envelope");
        const updateSend = config.action === "UpdateItem" && body.MessageDisposition === "SendAndSaveCopy";
        let sendMessage;
        if (updateSend) {
            const change = body.ItemChanges?.[0], id = change?.ItemId, shape = body.ItemShape;
            const fields = ["ToRecipients", "CcRecipients", "BccRecipients", "Subject", "Body"];
            if (config.body.__type !== "UpdateItemJsonRequest:#Exchange"
                || !onlyKeys(body, ["__type", "MessageDisposition", "ConflictResolution", "ClientSupportsIrm",
                    "SendCalendarInvitationsOrCancellations", "SuppressReadReceipts", "ComposeOperation",
                    "PromoteInlineAttachments", "SendOnNotFoundError", "ItemShape", "ItemChanges"])
                || body.ConflictResolution !== "NeverOverwrite" || body.ClientSupportsIrm !== true
                || body.SendCalendarInvitationsOrCancellations !== "SendToNone" || body.SuppressReadReceipts !== false
                || body.ComposeOperation !== "newMail" || body.PromoteInlineAttachments !== false
                || body.SendOnNotFoundError !== false
                || !onlyKeys(shape, ["__type", "BaseShape", "AdditionalProperties"])
                || shape.__type !== "ItemResponseShape:#Exchange" || shape.BaseShape !== "IdOnly"
                || !Array.isArray(shape.AdditionalProperties) || shape.AdditionalProperties.length !== 1
                || !onlyKeys(shape.AdditionalProperties[0], ["__type", "FieldURI"])
                || shape.AdditionalProperties[0].__type !== "PropertyUri:#Exchange"
                || shape.AdditionalProperties[0].FieldURI !== "ItemLastModifiedTime"
                || !Array.isArray(body.ItemChanges) || body.ItemChanges.length !== 1
                || !onlyKeys(change, ["__type", "ItemId", "Updates"]) || change.__type !== "ItemChange:#Exchange"
                || !onlyKeys(id, ["__type", "Id", "ChangeKey"]) || id.__type !== "ItemId:#Exchange"
                || typeof id.Id !== "string" || !id.Id.trim()
                || typeof id.ChangeKey !== "string" || !id.ChangeKey.trim()
                || !Array.isArray(change.Updates) || change.Updates.length !== fields.length) {
                throw new Error("Only one versioned verified draft submission is allowed");
            }
            sendMessage = {__type:"Message:#Exchange"};
            for (const update of change.Updates) {
                const field = update.Path?.FieldURI;
                if (!fields.includes(field) || Object.hasOwn(sendMessage, field)
                    || !onlyKeys(update, ["__type", "Path", "Item"]) || update.__type !== "SetItemField:#Exchange"
                    || !onlyKeys(update.Path, ["__type", "FieldURI"]) || update.Path.__type !== "PropertyUri:#Exchange"
                    || !onlyKeys(update.Item, ["__type", field]) || update.Item.__type !== "Message:#Exchange"
                    || !Object.hasOwn(update.Item, field)) {
                    throw new Error("Only exact recipient/content fields may be submitted");
                }
                sendMessage[field] = update.Item[field];
            }
        }
        if (config.action === "MoveItem") {
            const folder = body.ToFolderId?.BaseFolderId;
            if (body.ReturnNewItemIds !== true || folder?.__type !== "FolderId:#Exchange"
                || typeof folder.Id !== "string" || !folder.Id.trim()
                || !Array.isArray(body.ItemIds) || body.ItemIds.length !== 1
                || !body.ItemIds[0].Id || !body.ItemIds[0].ChangeKey) {
                throw new Error("Only a single versioned move to an explicit folder is allowed");
            }
        }
        if (config.action === "CreateItem" || updateSend) {
            const body = updateSend
                ? {MessageDisposition:"SendAndSaveCopy",Items:[sendMessage]}
                : config.body.Body;
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
            } else if (item.__type === "Message:#Exchange"
                && (body.MessageDisposition === "SendAndSaveCopy" || item.Attachments !== undefined)) {
                if (item.Attachments !== undefined) {
                    let total = 0;
                    const names = new Set();
                    const files = item.Attachments;
                    if (body.MessageDisposition !== "SaveOnly" || !Array.isArray(files) || !files.length || files.length > 5)
                        throw new Error("Attachments require a bounded unsent draft");
                    for (const file of files) {
                        if (!onlyKeys(file, ["__type", "Name", "ContentType", "IsInline", "IsContactPhoto", "Content"])
                            || file.__type !== "FileAttachment:#Exchange" || file.ContentType !== "application/octet-stream"
                            || file.IsInline !== false || file.IsContactPhoto !== false
                            || typeof file.Name !== "string" || !file.Name.trim() || [".", ".."].includes(file.Name)
                            || /[\x00-\x1f\x7f-\x9f/\\\u2028\u2029]/.test(file.Name)
                            || new TextEncoder().encode(file.Name).length > 255
                            || names.has(file.Name.toLowerCase()) || typeof file.Content !== "string"
                            || file.Content.length > 666668
                            || file.Content.length % 4 !== 0 || !/^[A-Za-z0-9+/]*={0,2}$/.test(file.Content)) {
                            throw new Error("Invalid bounded file attachment");
                        }
                        const bytes = atob(file.Content);
                        if (btoa(bytes) !== file.Content) throw new Error("Noncanonical attachment base64");
                        total += bytes.length;
                        names.add(file.Name.toLowerCase());
                    }
                    if (total > 500000) throw new Error("Attachments exceed 500000 bytes combined");
                }
                const addressOK = address => {
                    if (typeof address !== "string" || address.length > 254 || !/^[\x21-\x7e]+$/.test(address)) return false;
                    const parts = address.split("@");
                    return parts.length === 2 && parts[0].length <= 64
                        && /^[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+$/.test(parts[0])
                        && !parts[0].startsWith(".") && !parts[0].endsWith(".") && !parts[0].includes("..")
                        && parts[1].includes(".") && parts[1].split(".").every(label =>
                            label.length <= 63 && /^[A-Za-z0-9](?:[A-Za-z0-9-]*[A-Za-z0-9])?$/.test(label));
                };
                const lists = [item.ToRecipients, item.CcRecipients, item.BccRecipients];
                const recipients = lists.every(Array.isArray) ? lists.flat() : [];
                const text = item.Body;
                if (!onlyKeys(body, ["__type", "MessageDisposition", "Items"])
                    || !onlyKeys(item, ["__type", "Subject", "Body", "ToRecipients", "CcRecipients", "BccRecipients", "Attachments"])
                    || !lists.every(Array.isArray) || !item.ToRecipients.length || recipients.length > 20
                    || recipients.some(recipient => !onlyKeys(recipient, ["__type", "EmailAddress", "RoutingType"])
                        || recipient.__type !== "EmailAddress:#Exchange" || recipient.RoutingType !== "SMTP"
                        || !addressOK(recipient.EmailAddress))
                    || new Set(recipients.map(recipient => recipient.EmailAddress.toLowerCase())).size !== recipients.length
                    || typeof item.Subject !== "string" || !item.Subject.trim()
                    || /[\x00-\x1f\x7f-\x9f\u2028\u2029]/.test(item.Subject) || new TextEncoder().encode(item.Subject).length > 998
                    || !onlyKeys(text, ["__type", "BodyType", "Value"])
                    || text.__type !== "BodyContentType:#Exchange" || text.BodyType !== "Text"
                    || typeof text.Value !== "string" || !text.Value.trim() || text.Value.includes("\0")
                    || new TextEncoder().encode(text.Value).length > 200000) {
                    throw new Error("Only explicit-recipient plain-text mail is allowed; no sender override");
                }
            } else if (!["Message:#Exchange", "ReplyToItem:#Exchange", "ReplyAllToItem:#Exchange","Contact:#Exchange"].includes(item.__type)
                || body.MessageDisposition !== "SaveOnly") throw new Error("Sending is prohibited");
            if (item.__type === "Contact:#Exchange" && body.SavedItemFolderId?.BaseFolderId?.Id !== "contacts")
                throw new Error("Contacts must be created in the personal contacts folder");
        }
        if (config.action === "UpdateItem" && !updateSend && body.MessageDisposition !== "SaveOnly") throw new Error("Sending is prohibited");
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
