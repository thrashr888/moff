const {test} = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs"), path = require("node:path"), vm = require("node:vm");
const source = name => fs.readFileSync(path.join(__dirname, "../src", name), "utf8");
const json = (value, status = 200) => new Response(JSON.stringify(value), {status});
const itemUrl = "https://graph.microsoft.com/v1.0/me/drive/items/test";
const uploadUrl = "https://graph.microsoft.com/v1.0/me/drive/root:/synthetic.bin:/content?@microsoft.graph.conflictBehavior=fail";
const secretUrl = "https://tenant.sharepoint.com/transfer?secret=hidden";

function fixture(handler) {
    const requests = [], resources = [], window = {};
    const context = {location:{origin:"https://outlook.cloud.microsoft"}, URL, Uint8Array, window,
        setTimeout:(callback, delay) => { const timer = setTimeout(callback, delay); timer.unref(); return timer; },
        fetch:async (url, init) => {
            requests.push({url, init:{...init, body:init.body?.slice()}});
            return handler(url, init, requests.length);
        }};
    const adapter = vm.runInNewContext(`(${source("safari_transfer.js")})`, context);
    return {requests, resources, window, open:config => adapter(config, async resource => {
        resources.push(resource);
        return {status:200,accessToken:"test-secret"};
    }, async response => ({status:response.status,body:await response.text()}), new AbortController().signal)};
}

test("downloads split streams, preserve every byte and keep credentials inside the page", async () => {
    const data = Uint8Array.from({length:4000001}, (_,i) => i % 256);
    const f = fixture((url,init) => url === itemUrl
        ? json({file:{},size:data.length,"@microsoft.graph.downloadUrl":secretUrl})
        : new Response(new ReadableStream({start(controller) {
            const [,start,end] = /^bytes=(\d+)-(\d+)$/.exec(init.headers.Range);
            for (let i = Number(start); i <= Number(end); i += 65536) controller.enqueue(data.slice(i, Math.min(i + 65536,Number(end) + 1)));
            controller.close();
        }}),{status:206,headers:{"content-range":init.headers.Range.replace("=", " ") + "/" + data.length}}));
    const adapter = await f.open({url:itemUrl,size:data.length,direction:"download"});
    const chunks = [];
    while (true) {
        const bytes = await adapter.read(128000);
        if (bytes === null) break;
        assert.ok(bytes.length > 0 && bytes.length <= 128000);
        chunks.push(Buffer.from(bytes));
    }
    await adapter.close();
    assert.deepEqual(Buffer.concat(chunks), Buffer.from(data));
    assert.equal(f.requests[1].init.headers.authorization, undefined);
    assert.equal(f.requests.length,3);
    assert.equal(f.requests[1].init.credentials,"omit");
    assert.ok(!JSON.stringify(chunks).includes("hidden"));
});

test("download size mismatches, untrusted locations and oversized upstream chunks fail", async () => {
    for (const [size, declared, location, data] of [
        [2,3,secretUrl,[0,1]],
        [3,3,"https://evil.example/secret",[0,1,2]],
        [3,3,secretUrl,[0,1]],
        [3,3,secretUrl,[0,1,2,3]],
        [4000000,4000000,secretUrl,new Uint8Array(4000000)]
    ]) {
        const f = fixture(url => url === itemUrl ? json({file:{},size:declared,"@microsoft.graph.downloadUrl":location}) : new Response(new Uint8Array(data)));
        await assert.rejects(async () => {
            const adapter = await f.open({url:itemUrl,size,direction:"download"});
            while (await adapter.read(128000) !== null) {}
        });
    }
    const empty = fixture(url => url === itemUrl ? json({file:{},size:0,"@microsoft.graph.downloadUrl":secretUrl}) : new Response(new Uint8Array()));
    const adapter = await empty.open({url:itemUrl,size:0,direction:"download"});
    assert.equal(await adapter.read(0),null);
    await adapter.close();
});

test("SharePoint ranges work when CORS hides Content-Range, but mismatched exposed headers fail", async () => {
    for (const headers of [{}, {"content-length":"3"}, {"content-range":"bytes 0-2/3"}]) {
        const f = fixture(url => url === itemUrl ? json({file:{},size:3,"@microsoft.graph.downloadUrl":secretUrl})
            : new Response(new Uint8Array([0,128,255]),{status:206,headers}));
        const adapter = await f.open({url:itemUrl,size:3,direction:"download"});
        assert.deepEqual(Array.from(await adapter.read(128000)),[0,128,255]);
        assert.equal(await adapter.read(0),null);
        await adapter.close();
    }
    for (const headers of [{"content-length":"4"}, {"content-range":"bytes 1-3/4"}]) {
        const f = fixture(url => url === itemUrl ? json({file:{},size:3,"@microsoft.graph.downloadUrl":secretUrl})
            : new Response(new Uint8Array([0,128,255]),{status:206,headers}));
        await assert.rejects(f.open({url:itemUrl,size:3,direction:"download"}));
    }
});

test("uploads accumulate bounded Graph ranges independently of transport and do not export session URLs", async () => {
    const data = Uint8Array.from({length:4000001}, (_,i) => i % 256), ranges = [];
    const f = fixture((url, init) => {
        if (init.method === "POST") return json({uploadUrl:secretUrl});
        assert.equal(url,secretUrl);
        assert.equal(init.headers.authorization,undefined);
        const match = /^bytes (\d+)-(\d+)\/4000001$/.exec(init.headers["content-range"]);
        const end = Number(match[2]) + 1;
        ranges.push(Buffer.from(init.body));
        assert.ok(init.body.length <= 3276800);
        if (end < data.length) {
            assert.equal(init.body.length % 327680,0);
            return json({nextExpectedRanges:[`${end}-${data.length - 1}`]},202);
        }
        return json({id:"file-id",size:data.length,name:"synthetic.bin","@microsoft.graph.downloadUrl":secretUrl},201);
    });
    const adapter = await f.open({url:uploadUrl,direction:"upload",size:data.length,allow_mutation:true,result_key:"__moff_upload_test"});
    assert.equal(f.requests[0].url,"https://graph.microsoft.com/v1.0/me/drive/root:/synthetic.bin:/createUploadSession");
    assert.equal(JSON.parse(f.requests[0].init.body).item["@microsoft.graph.conflictBehavior"],"fail");
    for (let offset = 0; offset < data.length; offset += 128000) await adapter.write(data.slice(offset,offset + 128000),offset);
    await adapter.close();
    assert.deepEqual(Buffer.concat(ranges),Buffer.from(data));
    assert.equal(f.window.__moff_upload_test.id,"file-id");
    assert.ok(!JSON.stringify(f.window).includes("hidden"));
    assert.ok(!JSON.stringify(f.window).includes("test-secret"));
});

test("upload bounds, stale offsets, incomplete close and ambiguous ranges never trigger retries", async () => {
    const config = {url:uploadUrl,direction:"upload",size:3276801,allow_mutation:true,result_key:"__moff_upload_test"};
    const f = fixture((url, init) => init.method === "POST" ? json({uploadUrl:secretUrl}) : json({error:"unknown"},503));
    const adapter = await f.open(config);
    await assert.rejects(adapter.write(new Uint8Array(128001),0));
    await assert.rejects(adapter.write(new Uint8Array(1),1));
    await assert.rejects(adapter.close(),/Incomplete/);
    await assert.rejects(async () => {
        for (let offset = 0; offset < config.size; offset += 128000)
            await adapter.write(new Uint8Array(Math.min(128000,config.size-offset)),offset);
    });
    assert.equal(f.requests.length,2);
    assert.equal(f.window.__moff_upload_test_failure.stage,"upload range");
    assert.equal(f.window.__moff_upload_test_failure.status,503);
    assert.ok(!JSON.stringify(f.window).includes("hidden"));
    adapter.abort();
    assert.equal(f.window.__moff_upload_test,undefined);
    for (const overrides of [{size:250000001}, {allow_mutation:false}, {url:itemUrl}, {result_key:"unsafe"},
        {url:itemUrl + "/content",etag:null}]) {
        const denied = fixture(() => { throw new Error("must not fetch"); });
        await assert.rejects(denied.open({...config,...overrides}));
        assert.equal(denied.resources.length,0);
    }
});

test("existing uploads pin the entity tag and zero-byte uploads remain bounded", async () => {
    const f = fixture(() => json({id:"file-id",size:0},200));
    const adapter = await f.open({url:itemUrl + "/content",direction:"upload",size:0,
        allow_mutation:true,etag:"version",result_key:"__moff_upload_empty"});
    await adapter.close();
    assert.equal(f.requests[0].init.headers["if-match"],"version");
    assert.equal(f.requests[0].init.body.length,0);
    assert.equal(f.window.__moff_upload_empty.size,0);
});
