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
        121455: {
            ij: async () => options.expired ? {TokenResultCode: 3} : {TokenResultCode: 0, AccessToken: "test-secret"},
            a9: async () => options.expired ? "" : "synthetic-personal-token",
        },
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
        AbortController, setTimeout: options.setTimeout || setTimeout, clearTimeout, TextDecoder, TextEncoder,
        fetch: async (url, init) => {
            requests.push({url, init});
            if (options.networkFailure) throw new TypeError("synthetic network failure");
            if (options.hangFetch) return new Promise(() => {});
            return new Response(options.oversized ? "x".repeat(900001) : '{"Body":{"ResponseCode":"NoError"}}',
                {status: options.status || 200});
        },
    };
    const read = vm.runInNewContext(`(${source})`, context);
    const auth = vm.runInNewContext(`(${authSource})`, context);
    const response = vm.runInNewContext(`(${responseSource})`, context);
    return {call: config => read(config, options.acquireToken || auth, response), auth, requests, chunks};
}

function sendConfig() {
    const mailbox = address => ({__type:"EmailAddress:#Exchange",EmailAddress:address,RoutingType:"SMTP"});
    return {action:"CreateItem",mutation:true,allow_mutation:true,body:{
        Body:{__type:"CreateItemRequest:#Exchange",MessageDisposition:"SendAndSaveCopy",
            Items:[{__type:"Message:#Exchange",Subject:"Synthetic",
                Body:{__type:"BodyContentType:#Exchange",BodyType:"Text",Value:"<literal>\nText"},
                ToRecipients:[mailbox("to@example.com")],CcRecipients:[mailbox("cc@example.com")],
                BccRecipients:[mailbox("bcc@example.com")]}]}
    }};
}

test("typed explicit-recipient sending uses one same-origin request on work and personal Outlook", async () => {
    for (const origin of ["https://outlook.office.com","https://outlook.cloud.microsoft","https://outlook.live.com"]) {
        const f = fixture({origin});
        assert.equal((await f.call(sendConfig())).status,200);
        assert.equal(f.requests.length,1);
        assert.equal(f.requests[0].url,"/owa/service.svc?action=CreateItem&app=Mail");
        assert.equal(f.requests[0].init.redirect,"error");
        assert.deepEqual(JSON.parse(f.requests[0].init.body),sendConfig().body);
    }
});

test("browser send gate rejects implicit scope, spoofing, attachments and invalid content before auth or fetch", async () => {
    const changes = [
        config => {config.mutation=false;},
        config => {config.allow_mutation=false;},
        config => {config.action="SendItem";},
        config => {config.body.Body.MessageDisposition="SendOnly";},
        config => {config.body.Body.From={EmailAddress:"spoof@example.com"};},
        config => {config.body.Body.SavedItemFolderId={};},
        config => {config.body.Body.Items.push(config.body.Body.Items[0]);},
        ...["From","Sender","Attachments","ReferenceItemId","ItemId"].map(key =>
            config => {config.body.Body.Items[0][key]={};}),
        config => {config.body.Body.Items[0].ToRecipients=[];},
        config => {config.body.Body.Items[0].CcRecipients=null;},
        config => {delete config.body.Body.Items[0].BccRecipients;},
        config => {config.body.Body.Items[0].ToRecipients[0].RoutingType="EX";},
        config => {config.body.Body.Items[0].ToRecipients[0].Name="Hidden scope";},
        config => {config.body.Body.Items[0].CcRecipients[0].EmailAddress="TO@example.com";},
        ...["", "Sarah", "a@example.com,b@example.com", "a@example.com\r\nBcc:b@example.com",
            "a@@example.com", "a@localhost", ".a@example.com", "a..b@example.com",
            "a@-example.com", "a@example..com", " a@example.com", "ü@example.com"].map(address =>
            config => {config.body.Body.Items[0].ToRecipients[0].EmailAddress=address;}),
        config => {config.body.Body.Items[0].ToRecipients=Array.from({length:21},(_,i) =>
            ({__type:"EmailAddress:#Exchange",EmailAddress:`p${i}@example.com`,RoutingType:"SMTP"}));},
        ...["", " ", "Subject\nBcc: other@example.com", "Subject\u2028line", "Subject\u2029paragraph", "é".repeat(500)].map(subject =>
            config => {config.body.Body.Items[0].Subject=subject;}),
        ...["", " \n", "text\0", "é".repeat(100001)].map(body =>
            config => {config.body.Body.Items[0].Body.Value=body;}),
        config => {config.body.Body.Items[0].Body.BodyType="HTML";},
        config => {config.body.Body.Items[0].Body.Content="unsupported";},
        config => {config.body.Body.Items[0].__type="ReplyToItem:#Exchange";},
    ];
    for (const change of changes) {
        const config = sendConfig();
        change(config);
        let authentications=0;
        const f = fixture({acquireToken:async () => {authentications++; return {status:200};}});
        await assert.rejects(f.call(config));
        assert.equal(authentications,0);
        assert.equal(f.requests.length,0);
    }
});

test("send byte bounds permit the exact maximum and failures are never retried", async () => {
    const config = sendConfig();
    config.body.Body.Items[0].Subject="é".repeat(499);
    config.body.Body.Items[0].Body.Value="é".repeat(100000);
    config.body.Body.Items[0].ToRecipients=Array.from({length:18},(_,i) =>
        ({__type:"EmailAddress:#Exchange",EmailAddress:`p${i}@example.com`,RoutingType:"SMTP"}));
    assert.equal((await fixture().call(config)).status,200);
    config.body.Body.Items[0].BccRecipients.push(
        {__type:"EmailAddress:#Exchange",EmailAddress:"extra@example.com",RoutingType:"SMTP"});
    await assert.rejects(fixture().call(config));
    for (const options of [{networkFailure:true},{status:500},{hangFetch:true},{oversized:true}]) {
        const f = fixture(options);
        const result = await f.call({...sendConfig(),timeout_ms:5});
        assert.ok(result.failure || result.status === 500);
        assert.equal(f.requests.length,1);
    }
    const f = fixture({expired:true});
    assert.equal((await f.call(sendConfig())).status,401);
    assert.equal(f.requests.length,0);
});

test("archive move requires opt-in and exactly one versioned explicit-folder move", async () => {
    const body = {
        __type:"MoveItemRequest:#Exchange",
        ToFolderId:{__type:"TargetFolderId:#Exchange",BaseFolderId:{__type:"FolderId:#Exchange",Id:"archive-id"}},
        ItemIds:[{__type:"ItemId:#Exchange",Id:"message-id",ChangeKey:"key"}],
        ReturnNewItemIds:true,
    };
    const f = fixture();
    await f.call({action:"MoveItem",mutation:true,allow_mutation:true,body:{Body:body}});
    assert.equal(f.requests.length,1);
    for (const config of [
        {action:"MoveItem",body:{Body:body}},
        {action:"MoveItem",mutation:true,body:{Body:body}},
        {action:"MoveItem",mutation:true,allow_mutation:true,body:{Body:{...body,ItemIds:[]}}},
        {action:"MoveItem",mutation:true,allow_mutation:true,body:{Body:{...body,ItemIds:[{Id:"message-id"}]}}},
        {action:"MoveItem",mutation:true,allow_mutation:true,body:{Body:{...body,ToFolderId:{BaseFolderId:{__type:"DistinguishedFolderId:#Exchange",Id:"deleteditems"}}}}},
    ]) {
        const blocked = fixture();
        await assert.rejects(blocked.call(config));
        assert.equal(blocked.requests.length,0);
    }
});

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
    const drift = await fixture({changed: true}).call({action: "FindItem", body: {}});
    assert.equal(drift.status, 422);
    assert.equal(drift.stage, "authentication");
    const oversized = await fixture({oversized: true}).call({action: "FindItem", body: {}});
    assert.equal(oversized.failure, "response_too_large");
});
test("unloaded Outlook is distinct from expiry and adapter drift", async () => {
    const result = await fixture({unloaded:true}).call({action:"FindItem",body:{}});
    assert.equal(result.status,425);
    assert.equal(result.stage,"authentication");
});
test("network failure is not labelled session expiry and is never retried", async () => {
    const f = fixture({networkFailure:true});
    const result = await f.call({action:"FindItem",body:{}});
    assert.equal(result.failure,"network");
    assert.equal(result.dispatched,true);
    assert.equal(f.requests.length,1);
});
test("deadline covers stuck token acquisition and prevents a late mailbox dispatch", async () => {
    let finishAuth;
    const f = fixture({acquireToken:() => new Promise(resolve => {finishAuth=resolve;})});
    const result = await f.call({action:"FindItem",body:{},timeout_ms:1});
    assert.equal(result.failure,"timeout");
    assert.equal(result.dispatched,false);
    finishAuth({status:200,accessToken:"synthetic"});
    await new Promise(resolve => setImmediate(resolve));
    assert.equal(f.requests.length,0);
});
test("request deadline aborts stalled reads and records dispatched writes as uncertain", async () => {
    for (const config of [
        {action:"FindItem",body:{}},
        {action:"CreateItem",mutation:true,allow_mutation:true,
            body:{Body:{__type:"CreateItemRequest:#Exchange",MessageDisposition:"SaveOnly",
                Items:[{__type:"Message:#Exchange"}]}}},
    ]) {
        const f = fixture({hangFetch:true});
        const result = await f.call({...config,timeout_ms:1});
        assert.equal(result.failure,"timeout");
        assert.equal(result.dispatched,true);
        assert.equal(f.requests.length,1);
        assert.equal(f.requests[0].init.signal.aborted,true);
    }
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
