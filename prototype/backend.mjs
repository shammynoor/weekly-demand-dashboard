export class SupabaseBackend{
 constructor(client){this.client=client;this.demo=false;}
 async rpc(name,args={}){const {data,error}=await this.client.rpc('iq_'+name,args);if(error)throw error;return data;}
 async user(){const {data,error}=await this.client.auth.getUser();if(error)return null;return data.user;}
 async login(email,password){const {error}=await this.client.auth.signInWithPassword({email,password});if(error)throw error;}
 async signup(email,password,redirect){const {data,error}=await this.client.auth.signUp({email,password,options:{emailRedirectTo:redirect}});if(error)throw error;return !!data.session;}
 async logout(){const {error}=await this.client.auth.signOut({scope:'local'});if(error)throw error;}
 async reset(email,redirect){const {error}=await this.client.auth.resetPasswordForEmail(email,{redirectTo:redirect});if(error)throw error;}
 async password(password){const {error}=await this.client.auth.updateUser({password});if(error)throw error;}
 context(){return this.rpc('my_context');}
 claimPlatformAdmin(){return this.rpc('claim_platform_admin');}
 createPractice(name,tenant){return this.rpc('create_practice',{p_name:name,p_tenant_name:tenant});}
 async invite(practice,email,role){
  const {data,error}=await this.client.functions.invoke('iq-send-invite',{body:{practice,email,role}});
  if(error){let message='Invitation email could not be submitted. No successful send was confirmed.';try{const detail=await error.context?.json();if(detail?.error)message=detail.error;}catch{}throw Error(message);}
  if(!data?.email_submitted)throw Error('The email service did not confirm submission.');
  return data;
 }
 redeem(token){return this.rpc('redeem_invite',{p_token:token});}
 members(practice){return this.rpc('list_members',{p_practice:practice});}
 setStatus(practice,user,status){return this.rpc('set_member_status',{p_practice:practice,p_user:user,p_status:status});}
 importRH(practice,bundle,name){return this.rpc('import_rh',{p_practice:practice,p_tenant:bundle.tenant,p_source:name.slice(0,200),p_rows:bundle.rows});}
 async report(practice){
  const stateResult=await this.client.from('iq_rh_state').select('*').eq('practice_id',practice).maybeSingle();if(stateResult.error)throw stateResult.error;
  const daily=[];let from=0;
  while(true){
   const {data,error}=await this.client.from('iq_rh_daily').select('request_date,other_enquiries').eq('practice_id',practice).order('request_date').range(from,from+999);
   if(error)throw error;daily.push(...data);if(data.length<1000)break;from+=1000;
  }
  const result=await this.client.from('iq_rh_imports').select('source_name,imported_at,added,skipped').eq('practice_id',practice).order('imported_at',{ascending:false}).limit(8);if(result.error)throw result.error;
  return {state:stateResult.data,daily,imports:result.data};
 }
}
