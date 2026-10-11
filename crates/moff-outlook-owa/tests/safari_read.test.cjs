const {test} = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");
const source = fs.readFileSync(require("node:path").join(__dirname, "../src/safari_read.js"), "utf8");

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
        window: {webpackChunkOwa: chunks}, crypto: {randomUUID: () => "test"},
        AbortController, setTimeout, clearTimeout, TextDecoder,
        fetch: async (url, init) => {
            requests.push({url, init});
            return new Response(options.oversized ? "x".repeat(900001) : '{"Body":{"ResponseCode":"NoError"}}');
        },
    };
    return {call: vm.runInNewContext(`(${source})`, context), requests, chunks};
}

test("read authenticates inside page and returns no credential", async () => {
    const f = fixture();
    const result = await f.call({action: "FindItem", body: {}});
    assert.equal(result.status, 200);
    assert.equal(f.requests[0].init.headers.authorization, "Bearer test-secret");
    assert.equal(f.requests[0].init.redirect, "error");
    assert.equal(f.chunks.length, 0);
    assert.ok(!JSON.stringify(result).includes("test-secret"));
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
