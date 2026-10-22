const {test}=require("node:test");
const assert=require("node:assert/strict");
const vm=require("node:vm");
const fs=require("node:fs");
const path=require("node:path");
const source=fs.readFileSync(path.join(__dirname,"../src/safari_pin.js"),"utf8");
function fixture({changed=false,success=true}={}) {
    const calls=[];
    const modules={
        708334:{pinUnpinItems:async(...args)=>{calls.push(args);return {success};}},
        429896:{A:()=>({mailbox:"synthetic"})},
    };
    const runtime=id=>modules[id]; runtime.m=changed ? {} : modules;
    const chunks=[];
    chunks.push=entry=>{entry[2](runtime);Array.prototype.push.call(chunks,entry);};
    const context={location:{origin:"https://outlook.office.com"},window:{webpackChunkOwa:chunks},
        crypto:{randomUUID:()=>"unique"}};
    return {calls,context,call:vm.runInNewContext(`(${source})`,context)};
}
test("native pin calls only the observed helper with one exact message and no table mutation",async()=>{
    const f=fixture();
    assert.equal((await f.call({allow_mutation:true,id:"id",pinned:true})).success,true);
    assert.deepEqual(JSON.parse(JSON.stringify(f.calls)),[[["id"],{mailbox:"synthetic"},true]]);
});
test("unconfirmed or malformed pins never dispatch and adapter drift is explicit",async()=>{
    for(const args of [{id:"id",pinned:true},{allow_mutation:true,id:"",pinned:true},{allow_mutation:true,id:"id",pinned:"yes"}]) {
        const f=fixture();await assert.rejects(f.call(args));assert.equal(f.calls.length,0);
    }
    const f=fixture({changed:true});
    const result=await f.call({allow_mutation:true,id:"id",pinned:false});
    assert.equal(result.dispatched,false);
    assert.equal(result.reason,"auth_adapter_changed");
    assert.equal(f.calls.length,0);
});
test("unverified native acknowledgements never look like success and are not retried",async()=>{
    const f=fixture({success:false});
    const result=await f.call({allow_mutation:true,id:"id",pinned:false});
    assert.equal(result.success,false);assert.equal(result.dispatched,true);
    assert.equal(f.calls.length,1);
});
