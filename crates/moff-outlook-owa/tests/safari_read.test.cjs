const {test} = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");
const source = fs.readFileSync(require("node:path").join(__dirname, "../src/safari_read.js"), "utf8");
const authSource = fs.readFileSync(require("node:path").join(__dirname, "../src/safari_auth.js"), "utf8");
const responseSource = fs.readFileSync(require("node:path").join(__dirname, "../src/safari_response.js"), "utf8");

function fixture(options = {}) {
    const requests = [];
    const cache = {
        121455: {ij: async () => options.expired ? {TokenResultCode: 3} : {TokenResultCode: 0, AccessToken: "test-secret"}},
        429896: {A: () => ({})},
    };
    const runtime = id => cache[id];
    runtime.m = options.changed ? {} : cache;
    const chunks = [];
    chunks.push = function (entry) { entry[2](runtime); return Array.prototype.push.call(this, entry); };
    const context = {
        location: {origin: options.origin || "https://outlook.office.com"},
        window: {webpackChunkOwa: options.unloaded ? undefined : chunks}, crypto: {randomUUID: () => "test"},
        sessionStorage: {
            getItem: () => options.marker || null,
            removeItem: () => {},
        },
        AbortController, setTimeout, clearTimeout, TextDecoder,
        fetch: async (url, init) => {
            requests.push({url, init});
            return new Response(options.oversized ? "x".repeat(900001) : '{"Body":{"ResponseCode":"NoError"}}',
                {status: options.status || 200});
        },
    };
    const read = vm.runInNewContext(`(${source})`, context);
    const auth = vm.runInNewContext(`(${authSource})`, context);
    const response = vm.runInNewContext(`(${responseSource})`, context);
    return {call: config => read(config, auth, response), auth, requests, chunks};
}

test("authentication allowlist adds only the exact Substrate resource", async () => {
    const f = fixture();
    assert.equal((await f.auth("https://substrate.office.com")).status,200);
    assert.equal(f.chunks.length,0);
    for (const resource of ["http://substrate.office.com","https://substrate.office.com/","https://substrate.office.com.evil","https://evil.example"]) {
        await assert.rejects(f.auth(resource),/Unsupported/);
    }
});

test("read authenticates inside page and returns no credential", async () => {
    const f = fixture();
    const result = await f.call({action: "FindItem", body: {}});
    assert.equal(result.status, 200);
    assert.equal(f.requests[0].init.headers.authorization, "Bearer test-secret");
    assert.equal(f.requests[0].init.redirect, "error");
    assert.equal(f.chunks.length, 0);
    assert.ok(!JSON.stringify(result).includes("test-secret"));
});

test("SSO probe requires the recovery marker and exports only status, never credentials", async () => {
    for (const [options, status] of [
        [{}, 409],
        [{marker:"different"}, 409],
        [{marker:"recovery",expired:true}, 401],
        [{marker:"recovery"}, 200],
        [{marker:"recovery",unloaded:true}, 425],
    ]) {
        const f = fixture(options);
        const result = await f.call({action:"__MoffAuthStatus",recovery_marker:"recovery"});
        assert.deepEqual(JSON.parse(JSON.stringify(result)), {status});
        assert.equal(f.requests.length, 0);
        assert.ok(!JSON.stringify(result).includes("test-secret"));
    }
    const f = fixture({marker:"recovery",changed:true});
    assert.equal((await f.call({action:"__MoffAuthStatus",recovery_marker:"recovery"})).status, 422);
});

test("standalone auth status is silent and failed mailbox requests are not retried", async () => {
    const status = fixture();
    assert.equal((await status.call({action:"__MoffAuthStatus"})).status, 200);
    assert.equal(status.requests.length, 0);
    const failure = fixture({status:500});
    assert.equal((await failure.call({action:"FindItem",body:{}})).status, 500);
    assert.equal(failure.requests.length, 1);
});

test("refresh stores only a non-secret marker and schedules one reload without login UI interaction", () => {
    const refreshSource = fs.readFileSync(require("node:path").join(__dirname, "../src/safari_refresh.js"), "utf8");
    const stored = [], timers = [];
    let reloads = 0;
    const context = {
        location: {origin:"https://outlook.office.com", reload: () => { reloads++; }},
        sessionStorage: {setItem: (...args) => stored.push(args)},
        setTimeout: (callback, delay) => timers.push({callback, delay}),
    };
    const refresh = vm.runInNewContext(`(${refreshSource})`, context);
    assert.equal(refresh("marker").scheduled, true);
    assert.deepEqual(stored, [["__moff_auth_recovery", "marker"]]);
    assert.equal(reloads, 0);
    assert.equal(timers.length, 1);
    timers[0].callback();
    assert.equal(reloads, 1);
    context.location.origin = "https://login.microsoftonline.com";
    assert.throws(() => refresh("marker"), /HTTPS Outlook/);
    assert.equal(timers.length, 1);
});
test("rejects writes and foreign origins before fetching", async () => {
    for (const options of [{}, {origin: "https://example.com"}]) {
        const f = fixture(options);
        await assert.rejects(f.call({action: "DeleteItem", body: {}}));
        assert.equal(f.requests.length, 0);
    }
});
test("expired sessions do not attempt a mailbox request", async () => {
    const f = fixture({expired: true});
    assert.equal((await f.call({action: "FindItem", body: {}})).status, 401);
    assert.equal(f.requests.length, 0);
});
test("contract drift and oversized data fail explicitly", async () => {
    await assert.rejects(fixture({changed: true}).call({action: "FindItem", body: {}}), /contract changed/);
    await assert.rejects(fixture({oversized: true}).call({action: "FindItem", body: {}}), /exceeds/);
});
test("typed writes require gate, save-only disposition and supported actions", async () => {
    for (const config of [
        {action:"CreateItem",mutation:true,body:{Body:{__type:"CreateItemRequest:#Exchange",MessageDisposition:"SaveOnly"}}},
        {action:"SendItem",mutation:true,allow_mutation:true,body:{}},
        {action:"CreateItem",mutation:true,allow_mutation:true,body:{Body:{__type:"CreateItemRequest:#Exchange",MessageDisposition:"SendAndSaveCopy"}}},
        {action:"UpdateCalendarEvent",mutation:true,allow_mutation:true,body:{Body:{__type:"UpdateCalendarEventRequest:#Exchange",ShouldSendUpdateToAttendees:true}}}
    ]) {
        const f=fixture();
        await assert.rejects(f.call(config));
        assert.equal(f.requests.length,0);
    }
    const f=fixture();
    const result=await f.call({action:"CreateItem",mutation:true,allow_mutation:true,
        body:{Body:{__type:"CreateItemRequest:#Exchange",MessageDisposition:"SaveOnly",
            Items:[{__type:"ReplyToItem:#Exchange"}]}}});
    assert.equal(result.status,200);
    assert.ok(!JSON.stringify(result).includes("test-secret"));
    const contact = fixture();
    assert.equal((await contact.call({action:"CreateItem",mutation:true,allow_mutation:true,
        body:{Body:{__type:"CreateItemRequest:#Exchange",MessageDisposition:"SaveOnly",
            SavedItemFolderId:{BaseFolderId:{Id:"contacts"}},Items:[{__type:"Contact:#Exchange",GivenName:"Synthetic"}]}}})).status,200);
    await assert.rejects(fixture().call({action:"CreateItem",mutation:true,allow_mutation:true,
        body:{Body:{__type:"CreateItemRequest:#Exchange",MessageDisposition:"SaveOnly",
            SavedItemFolderId:{BaseFolderId:{Id:"inbox"}},Items:[{__type:"Contact:#Exchange"}]}}}));
});

test("meeting responses require a typed, versioned RSVP and explicit send disposition", async () => {
    const response = {Body:{__type:"CreateItemRequest:#Exchange",MessageDisposition:"SendAndSaveCopy",
        ItemShape:{__type:"ItemResponseShape:#Exchange",BaseShape:"IdOnly"},
        Items:[{__type:"DeclineItem:#Exchange",
            ReferenceItemId:{__type:"ItemId:#Exchange",Id:"synthetic",ChangeKey:"key"},
            Body:{__type:"BodyContentType:#Exchange",BodyType:"Text",Value:"I'll be at HashiConf."},
            ProposedStart:"2026-10-28T16:00:00Z",ProposedEnd:"2026-10-28T17:00:00Z"}]}};
    const f=fixture();
    assert.equal((await f.call({action:"CreateItem",mutation:true,allow_mutation:true,body:response})).status,200);
    for (const change of [
        {MessageDisposition:"SaveOnly"},
        {ToRecipients:[{EmailAddress:"other@example.com"}]},
        {Items:[{...response.Body.Items[0],ReferenceItemId:{Id:"synthetic"}}]},
        {Items:[{...response.Body.Items[0],Unexpected:"value"}]},
        {Items:[{...response.Body.Items[0],Body:{BodyType:"Text",Value:"message"}}]},
        {Items:[{...response.Body.Items[0],__type:"AcceptItem:#Exchange"}]},
        {Items:[{...response.Body.Items[0],ProposedEnd:undefined}]},
    ]) {
        const f=fixture();
        await assert.rejects(f.call({action:"CreateItem",mutation:true,allow_mutation:true,
            body:{Body:{...response.Body,...change}}}));
        assert.equal(f.requests.length,0);
    }
});

test("attendee notifications are allowed only for a versioned start/end move", async () => {
    const body = {Body:{__type:"UpdateCalendarEventRequest:#Exchange",
        EventId:{__type:"ItemId:#Exchange",Id:"synthetic",ChangeKey:"key"},
        EventScope:0,TargetAudience:0,ShouldSendUpdateToAttendees:true,
        ItemChange:{__type:"ItemChange:#Exchange",
            ItemId:{__type:"ItemId:#Exchange",Id:"synthetic",ChangeKey:"key"},
            Updates:[
                {__type:"SetItemField:#Exchange",Path:{__type:"PropertyUri:#Exchange",FieldURI:"Start"},
                    Item:{__type:"CalendarItem:#Exchange",Start:"2026-10-28T16:00:00Z"}},
                {__type:"SetItemField:#Exchange",Path:{__type:"PropertyUri:#Exchange",FieldURI:"End"},
                    Item:{__type:"CalendarItem:#Exchange",End:"2026-10-28T17:00:00Z"}}
            ]}}};
    const f=fixture();
    assert.equal((await f.call({action:"UpdateCalendarEvent",mutation:true,allow_mutation:true,body})).status,200);
    for (const change of [
        {EventScope:1},
        {ItemChange:{...body.Body.ItemChange,Updates:[...body.Body.ItemChange.Updates,
            {__type:"SetItemField:#Exchange",Path:{__type:"PropertyUri:#Exchange",FieldURI:"Subject"},
                Item:{__type:"CalendarItem:#Exchange",Subject:"Changed"}}]}},
    ]) {
        const f=fixture();
        await assert.rejects(f.call({action:"UpdateCalendarEvent",mutation:true,allow_mutation:true,
            body:{Body:{...body.Body,...change}}}));
        assert.equal(f.requests.length,0);
    }
});

test("recycle operations require one versioned item and prohibit permanent deletion or notifications", async () => {
    const body = {Body:{__type:"DeleteItemRequest:#Exchange",DeleteType:"MoveToDeletedItems",
        SendMeetingCancellations:"SendToNone",SuppressReadReceipts:true,
        ItemIds:[{Id:"synthetic",ChangeKey:"key"}]}};
    const f = fixture();
    assert.equal((await f.call({action:"DeleteItem",mutation:true,allow_mutation:true,body})).status,200);
    const cancel = fixture();
    assert.equal((await cancel.call({action:"DeleteItem",mutation:true,allow_mutation:true,
        body:{Body:{...body.Body,SendMeetingCancellations:"SendToAllAndSaveCopy"}}})).status,200);
    for (const change of [{DeleteType:"HardDelete"},{SendMeetingCancellations:"SendToAll"},
        {SuppressReadReceipts:false},{ItemIds:[{Id:"synthetic"}]}]) {
        const f = fixture();
        await assert.rejects(f.call({action:"DeleteItem",mutation:true,allow_mutation:true,body:{Body:{...body.Body,...change}}}));
        assert.equal(f.requests.length,0);
    }
});
