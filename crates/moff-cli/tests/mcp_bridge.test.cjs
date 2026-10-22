const {test}=require("node:test");
const assert=require("node:assert/strict");
const fs=require("node:fs");
const path=require("node:path");
const vm=require("node:vm");
const source=fs.readFileSync(path.join(__dirname,"../src/mcp_ui/bridge.js"),"utf8");
function fixture() {
    const sent=[],listeners={},timers=new Map(),nodes=new Map();
    let next=1;
    const parent={postMessage:m=>sent.push(m)};
    const document={
        documentElement:{dataset:{},scrollWidth:400},
        body:{getBoundingClientRect:()=>({height:300}),appendChild:n=>nodes.set(n.id,n)},
        addEventListener:(type,handler)=>listeners[type]=handler,
        getElementById:id=>nodes.get(id),
        createElement:()=>({setAttribute(){}}),
    };
    const window={parent,addEventListener:(type,handler)=>listeners[type]=handler};
    vm.runInNewContext(source,{window,document,navigator:{language:"en-US"},
        matchMedia:()=>({matches:false}),ResizeObserver:class{observe(){}},
        setTimeout:fn=>{const id=next++;timers.set(id,fn);return id;},clearTimeout:id=>timers.delete(id)});
    const respond=(id,result,error,from=parent)=>listeners.message({source:from,data:{jsonrpc:"2.0",id,result,error}});
    respond(sent[0].id,{hostContext:{theme:"dark"}});
    return {window,document,sent,listeners,timers,nodes,respond};
}
test("card tools call only their host and correlate exact structured results",async()=>{
    const f=fixture();
    const pending=f.window.moff.callTool("mail_get",{id:"synthetic"});
    const message=f.sent.at(-1);
    assert.equal(message.method,"tools/call");
    assert.deepEqual(JSON.parse(JSON.stringify(message.params)),{name:"mail_get",arguments:{id:"synthetic"}});
    f.respond(message.id,{isError:false,structuredContent:{ok:true,data:{body:"synthetic"}}});
    assert.equal((await pending).structuredContent.data.body,"synthetic");
});
test("foreign frames cannot answer requests or inject tool results",async()=>{
    const f=fixture();let rendered=0;f.window.render=()=>rendered++;
    const pending=f.window.moff.callTool("mail_get",{id:"synthetic"});
    const id=f.sent.at(-1).id;
    f.respond(id,{wrong:true},undefined,{});
    f.respond(id,{wrong:true},undefined,null);
    f.listeners.message({source:{},data:{jsonrpc:"2.0",method:"ui/notifications/tool-result",params:{}}});
    assert.equal(rendered,0);
    f.respond(id,{right:true});
    assert.equal((await pending).right,true);
});
test("Tavern native parent shims accept host-injected events without a source",async()=>{
    const f=fixture();
    f.window.webkit={messageHandlers:{mcp:{postMessage(){}}}};
    const pending=f.window.moff.callTool("mail_get",{id:"synthetic"});
    f.listeners.message({source:null,data:{jsonrpc:"2.0",id:f.sent.at(-1).id,result:{native:true}}});
    assert.equal((await pending).native,true);
});
test("host denial and missing acknowledgement reject without retrying",async()=>{
    const f=fixture();
    const denied=f.window.moff.callTool("mail_set_flag",{id:"synthetic",value:true});
    f.respond(f.sent.at(-1).id,undefined,{message:"Approval declined"});
    await assert.rejects(denied,e=>e.message==="Approval declined");
    const timed=f.window.moff.callTool("mail_set_pin",{id:"synthetic",value:true});
    [...f.timers.values()].at(-1)();
    await assert.rejects(timed,/Check the mailbox before retrying/);
    assert.equal(f.sent.filter(m=>m.method==="tools/call").length,2);
});
