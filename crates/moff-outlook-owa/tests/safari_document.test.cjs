const {test} = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs"), path = require("node:path"), vm = require("node:vm");
const source = fs.readFileSync(path.join(__dirname, "../src/documents/safari_document.js"), "utf8");
const base = "https://tenant.sharepoint.com/sites/Project";
const id = "12345678-1234-1234-1234-123456789abc";
const config = {base,id,size:3,etag:'"version-1"'};

function fixture(handler, origin = "https://tenant.sharepoint.com") {
    const requests = [], window = {};
    const adapter = vm.runInNewContext(`(${source})`, {
        URL,Uint8Array,location:{origin},window,
        setTimeout:(callback,delay) => {const timer=setTimeout(callback,delay);timer.unref();return timer;},
        fetch:async (url,init) => {requests.push({url,init}); return handler(url,init);}
    });
    return {requests,window,open:options => adapter({...config,...options},new AbortController().signal)};
}

test("Word metadata is exact-file scoped and never exports cookies or authentication", async () => {
    const f=fixture(()=>new Response(JSON.stringify({
        UniqueId:id.toUpperCase(),Name:"Notes.docx",Length:"3",ETag:'"version-1"',LockedByUser:null
    })));
    const value=await f.open({direction:"metadata"});
    assert.equal(value.id,id);
    assert.equal(value.etag,'"version-1"');
    assert.equal(value.locked,false);
    assert.equal(f.requests.length,1);
    assert.ok(f.requests[0].url.includes("GetFileById(guid'"+id+"')"));
    assert.equal(f.requests[0].init.credentials,"same-origin");
    assert.equal(f.requests[0].init.redirect,"error");
    assert.equal(f.requests[0].init.cache,"no-store");
    assert.equal(f.requests[0].init.headers.authorization,undefined);
    for (const corrupt of [
        {base:"https://evil.example/sites/Project"},
        {base:"https://tenant.sharepoint.com/sites/../Other"},
        {id:"../../files"},
        {id:id+"?extra"},
    ]) {
        const denied=fixture(()=>{throw Error("must not fetch");});
        await assert.rejects(denied.open({...corrupt,direction:"metadata"}));
        assert.equal(denied.requests.length,0);
    }
});

test("document streams preserve binary bytes and reject changed lengths", async () => {
    const f=fixture(()=>new Response(new Uint8Array([0,128,255]),{headers:{"content-length":"3"}}));
    const reader=await f.open({direction:"download"});
    assert.deepEqual(Array.from(await reader.read(2)),[0,128]);
    assert.deepEqual(Array.from(await reader.read(2)),[255]);
    assert.equal(await reader.read(0),null);
    await reader.close();
    assert.equal(f.requests[0].init.headers["if-match"],'"version-1"');
    const short=fixture(()=>new Response(new Uint8Array([0,128])));
    const incomplete=await short.open({direction:"download"});
    await incomplete.read(128000);
    await assert.rejects(incomplete.read(128000),/incomplete/);
});

test("document replacement is one atomic versioned write with page-owned digest and no replay", async () => {
    for (const status of [200,412,423]) {
        const f=fixture(url => url.endsWith("/contextinfo")
            ? new Response(JSON.stringify({FormDigestValue:"page-secret"}))
            : new Response("",{status}));
        const writer=await f.open({direction:"upload",allow_mutation:true,result_key:"__moff_doc_test"});
        await writer.write(new Uint8Array([0,128,255]),0);
        assert.equal(f.requests.length,0);
        if (status===200) await writer.close(); else await assert.rejects(writer.close(),/rejected/);
        const write=f.requests[1];
        assert.equal(write.init.method,"POST");
        assert.equal(write.init.headers["X-HTTP-Method"],"PUT");
        assert.equal(write.init.headers["If-Match"],'"version-1"');
        assert.equal(write.init.headers["X-RequestDigest"],"page-secret");
        assert.deepEqual(Array.from(write.init.body),[0,128,255]);
        assert.equal(f.window.__moff_doc_test.status,status);
        assert.ok(!JSON.stringify(f.window).includes("page-secret"));
        await assert.rejects(writer.close());
        assert.equal(f.requests.length,2);
    }
});

test("metadata reports editor locks without exporting the editor identity", async () => {
    const f=fixture(()=>new Response(JSON.stringify({
        UniqueId:id,Name:"Notes.docx",Length:"3",ETag:'"version-1"',LockedByUser:{Id:123}
    })));
    const value=await f.open({direction:"metadata"});
    assert.equal(value.locked,true);
    assert.equal(value.LockedByUser,undefined);
    assert.equal(value.Id,undefined);
    for (const lock of [undefined,null,{"odata.null":true}]) {
        const unlocked=fixture(()=>new Response(JSON.stringify({
            UniqueId:id,Name:"Notes.docx",Length:"3",ETag:'"version-1"',LockedByUser:lock
        })));
        assert.equal((await unlocked.open({direction:"metadata"})).locked,false);
    }
});

test("unapproved, wildcard, foreign, oversized and incomplete document writes dispatch nothing", async () => {
    for (const corrupt of [
        {allow_mutation:false},{etag:"*"},{etag:"bad\r\nheader"},{size:16000001},
        {result_key:"unsafe"},{base:"https://tenant.sharepoint.com.evil.example/sites/Project"},
    ]) {
        const f=fixture(()=>{throw Error("must not fetch");});
        await assert.rejects(f.open({direction:"upload",allow_mutation:true,result_key:"__moff_doc_test",...corrupt}));
        assert.equal(f.requests.length,0);
    }
    const f=fixture(()=>{throw Error("must not fetch");});
    const writer=await f.open({direction:"upload",allow_mutation:true,result_key:"__moff_doc_test"});
    await writer.write(new Uint8Array([1]),0);
    await assert.rejects(writer.close(),/Incomplete/);
    await writer.abort();
    assert.equal(f.requests.length,0);
});
