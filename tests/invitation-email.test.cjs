const {test}=require('node:test');
const assert=require('node:assert/strict');
const practice='11111111-1111-4111-8111-111111111111';
async function setup(options={}){
 const {makeHandler}=await import('../supabase/functions/iq-send-invite/handler.mjs');
 const calls=[];
 const audit={select(){return this;},eq(){return this;},gte(column){assert.equal(column,'occurred_at');return Promise.resolve({count:options.count||0});}};
 const handler=makeHandler({env:n=>({SUPABASE_URL:'https://example.supabase.co',SUPABASE_ANON_KEY:'public',BREVO_API_KEY:options.unconfigured?undefined:'private',INVITE_FROM_EMAIL:'sender@example.com'})[n],createClient:()=>({auth:{getUser:async()=>({data:{user:options.unauthenticated?null:{id:'user'}}})},from:()=>audit,rpc:async(name,args)=>{calls.push({name,args});return options.denied?{error:{message:'denied'}}:{data:{token:'secret-invite',expires_at:'2026-12-01'}};}}),fetch:async(url,request)=>{calls.push({url,request});return new Response(JSON.stringify(options.rejected?{message:'bad key'}:{messageId:'accepted'}),{status:options.rejected?400:201});}});
 const invoke=(body={},headers={},method='POST')=>handler(new Request('https://function.example',{method,headers:{authorization:'Bearer session',origin:'https://darwin-iq-practice-prototype.shammy-noor.chatgpt.site',...headers},...(method==='POST'?{body:JSON.stringify({practice,email:' Person@Example.com ',role:'viewer',...body})}:{})}));
 return {calls,invoke};
}
test('sends only after checked RPC, to normalized recipient, with server-owned link',async()=>{
 const {calls,invoke}=await setup();const response=await invoke({redirect:'https://attacker.example'});assert.equal(response.status,200);assert.equal(calls[0].name,'iq_create_invite');assert.equal(calls[0].args.p_email,'person@example.com');const mail=JSON.parse(calls[1].request.body);assert.deepEqual(mail.to,[{email:'person@example.com'}]);assert.match(mail.textContent,/https:\/\/darwin-iq-practice-prototype.shammy-noor.chatgpt.site\/\?invite=secret-invite/);assert.doesNotMatch(mail.textContent,/attacker/);const result=await response.json();assert.equal(result.email_submitted,true);assert.equal(result.token,undefined);
});
test('unauthenticated, wrong-origin, invalid-role and malformed-practice requests never send',async()=>{
 for(const [config,body,headers,status] of [[{unauthenticated:true},{},{},401],[{},{},{origin:'https://attacker.example'},403],[{},{role:'platform_admin'},{},400],[{},{practice:'wrong'},{},400]]){const {calls,invoke}=await setup(config);assert.equal((await invoke(body,headers)).status,status);assert.equal(calls.length,0);}
});
test('unauthorized practice and rate-limited administrator never send',async()=>{
 for(const config of [{denied:true},{count:20}]){const {calls,invoke}=await setup(config);assert.ok([403,429].includes((await invoke()).status));assert.equal(calls.filter(c=>c.url).length,0);}
});
test('missing configuration creates no invitation; provider rejection is never success',async()=>{
 let state=await setup({unconfigured:true});assert.equal((await state.invoke()).status,503);assert.equal(state.calls.length,0);state=await setup({rejected:true});const response=await state.invoke();assert.equal(response.status,502);assert.equal((await response.json()).email_submitted,undefined);
});
