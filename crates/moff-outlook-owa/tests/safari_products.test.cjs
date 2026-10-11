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
            return new Response(options.oversized ? "x".repeat(900001) : '{"value":[{"id":"test"}]}',
                {status: options.status || 200});
        },
    };
    const read = vm.runInNewContext(`(${source("safari_products.js")})`, context);
    const response = vm.runInNewContext(`(${source("safari_response.js")})`, context);
    return {
        requests, resources,
        call: url => read({url}, async resource => {
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
