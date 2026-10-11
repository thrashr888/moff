const {test} = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs"), path = require("node:path"), vm = require("node:vm");
const source = name => fs.readFileSync(path.join(__dirname, "../src", name), "utf8");

function fixture(options = {}) {
    const requests = [], resources = [];
    const context = {
        location: {origin: options.origin || "https://outlook.cloud.microsoft"},
        URL, AbortController, TextDecoder, setTimeout, clearTimeout,
        fetch: async (url, init) => {
            requests.push({url, init});
            if (requests.length === 2 && options.downloadBytes) return new Response(new Uint8Array(options.downloadBytes));
            return new Response(options.status === 204 ? null : options.body || (options.oversized ? "x".repeat(900001) : '{"value":[{"id":"test"}]}'),
                {status: options.status || 200});
        },
    };
    const read = vm.runInNewContext(`(${source("safari_products.js")})`, context);
    const response = vm.runInNewContext(`(${source("safari_response.js")})`, context);
    return {
        requests, resources,
        call: (url, config = {}) => read({url, ...config}, async resource => {
            resources.push(resource);
            return {status: options.authStatus || 200, accessToken: "test-secret"};
        }, response),
    };
}

test("product metadata uses a Graph token only inside the page", async () => {
    const f = fixture();
    const result = await f.call("https://graph.microsoft.com/v1.0/me/people?$top=1");
    assert.equal(result.status, 200);
    assert.deepEqual(f.resources, ["https://graph.microsoft.com"]);
    assert.equal(f.requests.length, 1);
    assert.equal(f.requests[0].init.method, "GET");
    assert.equal(f.requests[0].init.credentials, "omit");
    assert.equal(f.requests[0].init.mode, "cors");
    assert.equal(f.requests[0].init.redirect, "error");
    assert.ok(!JSON.stringify(result).includes("test-secret"));
});

test("typed contact, folder and group writes stay in-page and handle bodyless success", async () => {
    for (const [suffix, method, body, etag] of [
        ["/me/contacts","POST",{givenName:"Synthetic"}],
        ["/me/contacts/id","PATCH",{surname:"Updated"}],
        ["/me/drive/root/children","POST",{name:"Synthetic",folder:{},"@microsoft.graph.conflictBehavior":"fail"}],
        ["/me/drive/items/id","PATCH",{name:"Updated"},"etag"],
        ["/groups","POST",{displayName:"Test",mailNickname:"test",groupTypes:["Unified"],mailEnabled:true,securityEnabled:false}],
        ["/groups/id","PATCH",{description:"Test"}],
        ["/groups/id/members/$ref","POST",{"@odata.id":"https://graph.microsoft.com/v1.0/users/user-id"}],
    ]) {
        const f = fixture();
        await f.call("https://graph.microsoft.com/v1.0" + suffix, {method,body,etag,allow_mutation:true});
        assert.equal(f.requests[0].init.method, method);
        assert.deepEqual(JSON.parse(f.requests[0].init.body), body);
    }
    const f = fixture({status:204});
    const result = await f.call("https://graph.microsoft.com/v1.0/me/drive/items/id", {method:"DELETE",etag:"etag",allow_mutation:true});
    assert.equal(result.status,204);
    assert.equal(f.requests[0].init.headers["if-match"],"etag");
});

test("invalid and unconfirmed product writes never authenticate or fetch", async () => {
    for (const [suffix, config] of [
        ["/me/contacts",{method:"POST",body:{givenName:"Test"}}],
        ["/users/id",{method:"DELETE",allow_mutation:true}],
        ["/me/contacts/id",{method:"PATCH",body:{owners:[]},allow_mutation:true}],
        ["/me/drive/items/id",{method:"PATCH",body:{name:"new"},allow_mutation:true}],
        ["/me/drive/items/id",{method:"DELETE",etag:"bad\r\nheader",allow_mutation:true}],
        ["/me/drive/root/children",{method:"POST",body:{name:"bad/name",folder:{},"@microsoft.graph.conflictBehavior":"fail"},allow_mutation:true}],
        ["/groups",{method:"POST",body:{displayName:"Test",mailNickname:"test",groupTypes:["Security"],mailEnabled:false,securityEnabled:true},allow_mutation:true}],
        ["/groups/id/members/$ref",{method:"POST",body:{"@odata.id":"https://evil.example/v1.0/users/user"},allow_mutation:true}],
        ["/groups/id/members/$ref",{method:"POST",body:{"@odata.id":"https://graph.microsoft.com/v1.0/groups/id"},allow_mutation:true}],
        ["/groups/id/members/user",{method:"DELETE",allow_mutation:true}],
        ["/me/contacts",{method:"GET",body:{givenName:"Test"}}],
    ]) {
        const f = fixture();
        await assert.rejects(f.call("https://graph.microsoft.com/v1.0"+suffix,config));
        assert.equal(f.resources.length,0);
        assert.equal(f.requests.length,0);
    }
});

test("preauthenticated download URLs never leave the browser response", async () => {
    const f = fixture({body:JSON.stringify({value:[{id:"test","@microsoft.graph.downloadUrl":"https://download.example/secret","@microsoft.graph.downloadUrlNoAuth":"https://download.example/other"}]})});
    const result = await f.call("https://graph.microsoft.com/v1.0/me/drive/root/children");
    assert.ok(!result.body.includes("downloadUrl"));
    assert.ok(!result.body.includes("secret"));
});

test("single-shot transfers are refused in favor of the streamed adapter", async () => {
    await assert.rejects(fixture().call("https://graph.microsoft.com/v1.0/me/drive/items/test",{download:true}));
    await assert.rejects(fixture().call("https://graph.microsoft.com/v1.0/me/drive/items/test/content",
        {method:"PUT",bytes:[256],etag:"etag",allow_mutation:true}));
    await assert.rejects(fixture().call("https://graph.microsoft.com/v1.0/me/drive/root:/test.bin:/content",
        {method:"PUT",bytes:[0],allow_mutation:true}));
    const member = fixture({status:204});
    await member.call("https://graph.microsoft.com/v1.0/groups/group-id/members/user-id/$ref",{method:"DELETE",allow_mutation:true});
    assert.equal(member.requests[0].init.method,"DELETE");
    assert.equal(member.requests[0].init.body,undefined);
});

test("non-product routes and unsafe queries fail before auth or fetch", async () => {
    for (const suffix of [
        "/v1.0/me/messages", "/v1.0/me/events", "/v1.0/users/id/messages",
        "/v1.0/me/drive/items/id/content", "/v1.0/me/drive/items/%2Fmessages",
        "/v1.0/me/people?$top=101", "/v1.0/me/people?$top=1&$top=2",
        "/v1.0/me/people?$expand=messages", "/v1.0/users/id?$select=password",
        "/v1.0/groups?$filter=anything",
        "/v1.0/me/drive/items/%2e%2e/root/children",
        "/v1.0/me//people", "/v1.0/me/people/",
        "/v1.0/me/people?$top=%2B1",
    ]) {
        const f = fixture();
        await assert.rejects(f.call("https://graph.microsoft.com" + suffix));
        assert.equal(f.resources.length, 0);
        assert.equal(f.requests.length, 0);
    }
    for (const url of [
        "https://evil.example/v1.0/me/people", "http://graph.microsoft.com/v1.0/me/people",
        "https://user@graph.microsoft.com/v1.0/me/people", "https://graph.microsoft.com/v1.0/me/people#fragment",
    ]) {
        await assert.rejects(fixture().call(url));
    }
});

test("expired product auth does not fetch and permissions never trigger retries", async () => {
    const expired = fixture({authStatus:401});
    assert.equal((await expired.call("https://graph.microsoft.com/v1.0/groups")).status, 401);
    assert.equal(expired.requests.length, 0);
    const denied = fixture({status:403});
    assert.equal((await denied.call("https://graph.microsoft.com/v1.0/groups")).status, 403);
    assert.equal(denied.requests.length, 1);
});

test("foreign browser origins, unavailable contracts and oversized responses fail explicitly", async () => {
    await assert.rejects(fixture({origin:"https://evil.example"}).call("https://graph.microsoft.com/v1.0/groups"));
    await assert.rejects(fixture({authStatus:422}).call("https://graph.microsoft.com/v1.0/groups"), /contract/);
    await assert.rejects(fixture({oversized:true}).call("https://graph.microsoft.com/v1.0/groups"), /exceeds/);
});
