const {test} = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs"), path = require("node:path"), vm = require("node:vm");
const {randomUUID} = require("node:crypto");
const source = name => fs.readFileSync(path.join(__dirname, "../src", name), "utf8");
const base = "https://substrate.office.com/recommended/api/v1.1/loop";
const claims = {oid:"00000000-0000-0000-0000-000000000001",tid:"00000000-0000-0000-0000-000000000002",upn:"owner@example.com"};
const token = `header.${Buffer.from(JSON.stringify(claims)).toString("base64url")}.test-secret`;
const pod = btoa("ODSP|tenant.sharepoint.com|b!drive|root");
const workspace = {id:"SPO_test",title:"Project",mfs_info:{pod_id:pod},permissions:[{role:"Owner",upn:"OWNER@example.com"}]};

function fixture(handler) {
    const requests = [], resources = [];
    const context = {location:{origin:"https://outlook.cloud.microsoft"}, URL, atob, btoa,
        crypto:{randomUUID}, AbortController, setTimeout, clearTimeout,
        fetch:async (url, init) => {
            requests.push({url,init});
            const [value,status=200] = handler(url,init,requests.length);
            return new Response(JSON.stringify(value),{status});
        }};
    const adapter = vm.runInNewContext(`(${source("safari_loop.js")})`,context);
    return {requests,resources,call:config => adapter(config,async resource => {
        resources.push(resource);
        return {status:200,accessToken:token};
    },async response => ({status:response.status,body:await response.text()}))};
}

test("recent workspaces use PodIds, title search is literal and incompleteness is explicit", async () => {
    const f = fixture(() => [{workspaces:[workspace,{...workspace,title:"Other"}],pages:[],is_complete:true}]);
    const result = await f.call({action:"list",query:"PROJECT",limit:1});
    const value = JSON.parse(result.body);
    assert.equal(value.workspaces.length,1);
    assert.equal(value.workspaces[0].id,pod);
    assert.equal(value.workspaces[0].source_id,"SPO_test");
    assert.equal(value.is_complete,true);
    assert.equal(f.requests[0].url,base + "/recent");
    assert.equal(f.requests[0].init.headers["X-AnchorMailbox"],`Oid:${claims.oid}@${claims.tid}`);
    assert.match(f.requests[0].init.headers["Client-Request-Id"],/^[a-f0-9-]{36}$/);
    assert.ok(!JSON.stringify(result).includes("test-secret"));
    const limited = await f.call({action:"list",limit:1});
    assert.equal(JSON.parse(limited.body).is_complete,false);
});

test("workspace creation derives ODSP PodId from site domain, drive and Graph root", async () => {
    const f = fixture((url,init,index) => index === 1
        ? [{workspace:{id:"b!drive",sharepoint_info:{site_url:"https://TENANT.sharepoint.com/sites/test"}}},201]
        : [{id:"root"}]);
    const result = JSON.parse((await f.call({action:"create",allow_mutation:true})).body);
    assert.equal(result.id,pod);
    assert.equal(f.requests[0].url,base + "/workspaces");
    assert.deepEqual(JSON.parse(f.requests[0].init.body),{type:"tenantOwned"});
    assert.equal(f.requests[1].url,"https://graph.microsoft.com/v1.0/drives/b!drive/root?$select=id");
    assert.deepEqual(f.resources,["https://substrate.office.com","https://graph.microsoft.com"]);
    assert.equal(result.title_supported,false);
    const failed = fixture((url,init,index) => index === 1
        ? [{workspace:{id:"b!drive",sharepoint_info:{site_url:"https://tenant.sharepoint.com/sites/test"}}},201]
        : [{error:"denied"},403]);
    const partial = JSON.parse((await failed.call({action:"create",allow_mutation:true})).body);
    assert.equal(partial.drive_id,"b!drive");
    assert.equal(partial.verification_failed,true);
    assert.equal(failed.requests.length,2);
});

test("permanent deletion checks current-user Owner, includes MRU when known, then verifies 404", async () => {
    const f = fixture((url,init,index) => index === 1 ? [{workspaces:[workspace]}]
        : index === 2 ? [{response_code:10010}] : [{error:{code:"itemNotFound"}},404]);
    const mru = "00000000-0000-0000-0000-000000000003";
    const result = JSON.parse((await f.call({action:"delete",id:pod,mru_id:mru,allow_mutation:true})).body);
    assert.equal(f.requests[0].url,base + "/workspaces/" + encodeURIComponent(pod));
    assert.equal(f.requests[1].init.method,"DELETE");
    assert.deepEqual(JSON.parse(f.requests[1].init.body),{mru_id:mru});
    assert.equal(result.permanent,true);
    assert.equal(result.verified,true);
    for (const permissions of [[{role:"Member",upn:claims.upn}],[{role:"Owner",upn:"someone@example.com"}],[]]) {
        const denied = fixture(() => [{workspaces:[{...workspace,permissions}]}]);
        assert.equal((await denied.call({action:"delete",id:pod,allow_mutation:true})).status,403);
        assert.equal(denied.requests.length,1);
    }
});

test("unsafe IDs, unconfirmed writes and unsupported update fail before authentication", async () => {
    for (const config of [{action:"get",id:"SPO_test"},{action:"get",id:"b!drive"},
        {action:"get",id:btoa("ODSP|tenant.sharepoint.com|b!drive|root|extra")},
        {action:"create"},{action:"delete",id:pod},{action:"update",id:pod,allow_mutation:true},
        {action:"list",limit:101},{action:"list",limit:10,query:""}]) {
        const f = fixture(() => { throw new Error("must not fetch"); });
        await assert.rejects(f.call(config));
        assert.equal(f.resources.length,0);
    }
});

test("failed delete acknowledgement or read-back stops without retries", async () => {
    for (const code of [2,10010]) {
        const f = fixture((url,init,index) => index === 1 ? [{workspaces:[workspace]}]
            : index === 2 ? [{response_code:code}] : [{workspaces:[workspace]}]);
        await assert.rejects(f.call({action:"delete",id:pod,allow_mutation:true}));
        assert.equal(f.requests.length,code === 2 ? 2 : 3);
        assert.deepEqual(JSON.parse(f.requests[1].init.body),{});
    }
});
