const {test} = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");
const path = require("node:path");
const html = fs.readFileSync(path.join(__dirname, "../src/mcp_ui/inbox.html"), "utf8");
const script = [...html.matchAll(/<script>([\s\S]*?)<\/script>/g)].at(-1)[1];

function fixture(data, calls = async () => { throw new Error("Unexpected tool call"); }) {
    const app = {};
    const listeners = {};
    const context = vm.createContext({
        document: {getElementById: () => app, addEventListener:(type,handler)=>{listeners[type]=handler;}},
        window:{moff:{callTool:calls}},
    });
    vm.runInContext(script, context);
    context.render({structuredContent: {ok:true, data}}, {});
    return {app, context, listeners};
}
function render(data) { const {app}=fixture(data); return app.textContent || app.innerHTML; }

function mail(id,conversation,received="2026-10-21T12:00:00Z") {
    return {id,conversation_id:conversation,received,change_key:"key",subject:"Synthetic subject",
        from:{name:"Synthetic sender"},is_draft:false,is_read:false,flagged:false,pinned:false};
}
const success = data => ({structuredContent:{ok:true,data}});

test("mail actions use icons with tooltips and accessible stateful labels", async () => {
    const {app,context}=fixture({folder:"inbox",mail_actions_enabled:true,
        messages:[mail("a","thread"),mail("b","thread")]},async()=>success({
            verified:true,value:true,change_key:"updated",
        }));
    for (const label of ["Flag latest","Pin latest","Archive 2 loaded"]) {
        const button=app.innerHTML.match(new RegExp(`<button[^>]*title="${label}"[^>]*>[\\s\\S]*?<\\/button>`))?.[0];
        assert.ok(button,`Missing tooltip: ${label}`);
        assert.match(button,new RegExp(`aria-label="${label}"`));
        assert.match(button,/<svg[^>]*aria-hidden="true"/);
        assert.doesNotMatch(button,new RegExp(`>${label}<`));
    }
    assert.match(app.innerHTML,/>Refresh<\/button>/);
    await context.act("flag","conversation:thread");
    assert.match(app.innerHTML,/title="Unflag latest" aria-label="Unflag latest" aria-pressed="true"/);
    await context.act("pin","conversation:thread");
    assert.match(app.innerHTML,/title="Unpin latest" aria-label="Unpin latest" aria-pressed="true"/);
});

test("empty search pages show incomplete scope and the continuation offset", () => {
    const text = render({query:"synthetic", messages:[], scanned:20, has_more:true, next_offset:40});
    assert.match(text, /No matches in this page/);
    assert.match(text, /Search incomplete; continue at offset 40/);
    assert.doesNotMatch(text, /No matching mail\./);
});

test("both nonempty and exhausted search pages disclose scan progress", () => {
    assert.match(render({query:"synthetic",messages:[],scanned:3,has_more:false}), /Reached end of folder/);
    assert.match(render({query:"synthetic",folder:"inbox",messages:[{subject:"test"}],
        scanned:20,has_more:true,next_offset:20}), /continue at offset 20/);
});

test("conversation IDs group messages, latest wins, and same subjects alone never merge", () => {
    const {app,context}=fixture({folder:"inbox",messages:[
        mail("old","thread"), mail("new","thread","2026-10-21T13:00:00Z"),
        mail("different","other"),mail("no-conversation",null),mail("new","thread"),
    ]});
    const groups=context.groupMessages([
        mail("old","thread"),mail("new","thread","2026-10-21T13:00:00Z"),mail("different","other")
    ]);
    assert.equal(groups.length,2);
    assert.equal(groups[0].latest.id,"new");
    assert.match(app.innerHTML,/3 conversations/);
    assert.match(app.innerHTML,/2 loaded/);
    assert.equal((app.innerHTML.match(/<details /g)||[]).length,3);
    assert.doesNotMatch(app.innerHTML,/<details[^>]* open/);
});

test("expansion loads only the latest message, escapes bodies and keeps collapse state", async () => {
    const calls=[];
    const {context,app,listeners}=fixture({folder:"inbox",messages:[mail("id","thread")]},async(name,args)=>{
        calls.push({name,args});return success({id:"id",body:"<script>untrusted</script>"});
    });
    listeners.toggle({target:{dataset:{thread:"conversation:thread"},open:true}});
    await new Promise(resolve=>setImmediate(resolve));
    assert.deepEqual(JSON.parse(JSON.stringify(calls)),[{name:"mail_get",args:{id:"id"}}]);
    assert.match(app.innerHTML,/<details[^>]* open/);
    assert.match(app.innerHTML,/&lt;script&gt;untrusted/);
    assert.doesNotMatch(app.innerHTML,/<script>untrusted/);
    listeners.toggle({target:{dataset:{thread:"conversation:thread"},open:false}});
    vm.runInContext("draw()",context);
    assert.doesNotMatch(app.innerHTML,/<details[^>]* open/);
});

test("flag and pin send explicit desired state and update change keys only after verification", async () => {
    const calls=[];
    const {context,app}=fixture({folder:"inbox",mail_actions_enabled:true,messages:[mail("id","thread")]},async(name,args)=>{
        calls.push({name,args});return success({verified:true,value:true,change_key:"new-key"});
    });
    await context.act("flag","conversation:thread");
    await context.act("pin","conversation:thread");
    assert.equal(calls[0].name,"mail_set_flag");
    assert.equal(calls[0].args.value,true);
    assert.equal(calls[1].name,"mail_set_pin");
    assert.equal(calls[1].args.change_key,"new-key");
    assert.match(app.innerHTML,/Unflag latest/);
    assert.match(app.innerHTML,/Unpin latest/);
});

test("archive previews only loaded IDs, waits for confirmation and invalidates paging", async () => {
    const calls=[],plan={messages:[{id:"a",subject:"one"},{id:"b",subject:"two"}]};
    const {context,app}=fixture({folder:"inbox",mail_actions_enabled:true,has_more:true,next_offset:20,
        messages:[mail("a","thread"),mail("b","thread")]},async(name,args)=>{
        calls.push({name,args});
        return success(name==="mail_archive_preview" ? plan : {ok:true,completed:[]});
    });
    await context.act("archive","conversation:thread");
    assert.equal(calls.length,1);
    assert.equal(calls[0].name,"mail_archive_preview");
    assert.deepEqual(Array.from(calls[0].args.ids),["a","b"]);
    assert.match(app.innerHTML,/Continue to approval/);
    await context.act("confirm-archive","");
    assert.equal(calls[1].name,"mail_archive");
    assert.deepEqual(JSON.parse(JSON.stringify(calls[1].args.plan)),plan);
    assert.match(app.innerHTML,/Inbox changed/);
    assert.doesNotMatch(app.innerHTML,/<details /);
    await context.act("more","");
    assert.equal(calls.length,2);
});

test("denied and uncertain actions do not claim success or retry", async () => {
    let calls=0;
    const {context,app}=fixture({folder:"inbox",mail_actions_enabled:true,messages:[mail("id","thread")]},
        async()=>{calls++;return {isError:true,structuredContent:{ok:false,error:{message:"Approval declined"}}};});
    await context.act("flag","conversation:thread");
    assert.equal(calls,1);
    assert.match(app.innerHTML,/Approval declined/);
    assert.doesNotMatch(app.innerHTML,/Unflag latest/);
});
