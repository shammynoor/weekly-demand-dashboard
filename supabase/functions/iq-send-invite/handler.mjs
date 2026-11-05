const SITE = 'https://darwin-iq-practice-prototype.shammy-noor.chatgpt.site';
const emailPattern = /^[^@\s]+@[^@\s]+\.[^@\s]+$/;
const uuidPattern = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

// No service-role key: invitation creation uses the caller's checked database RPC.
export function makeHandler({env, createClient, fetch: send}) {
 return async request => {
  const headers = {'Access-Control-Allow-Origin': SITE, 'Access-Control-Allow-Headers': 'authorization, apikey, content-type, x-client-info', 'Access-Control-Allow-Methods': 'POST, OPTIONS', 'Content-Type': 'application/json', 'Vary': 'Origin'};
  const reply = (status, body) => new Response(JSON.stringify(body), {status, headers});
  if (request.headers.get('origin') && request.headers.get('origin') !== SITE) return reply(403, {error:'Origin not allowed.'});
  if (request.method === 'OPTIONS') return new Response(null, {status:204, headers});
  if (request.method !== 'POST') return reply(405, {error:'POST required.'});
  const authorization = request.headers.get('authorization');
  if (!authorization?.startsWith('Bearer ')) return reply(401, {error:'Sign in first.'});
  try {
   const client = createClient(env('SUPABASE_URL'), env('SUPABASE_ANON_KEY'), {global:{headers:{Authorization:authorization}}, auth:{persistSession:false,autoRefreshToken:false}});
   const {data:identity,error:authError} = await client.auth.getUser(authorization.slice(7));
   if (authError || !identity.user) return reply(401, {error:'Sign in again before sending an invitation.'});
   const body = await request.text();
   if (body.length > 2048) return reply(400, {error:'Invitation request too large.'});
   let input;
   try { input = JSON.parse(body); } catch { return reply(400, {error:'Invalid invitation request.'}); }
   const email = typeof input?.email === 'string' ? input.email.trim().toLowerCase() : '';
   if (!uuidPattern.test(input?.practice || '') || !emailPattern.test(email) || email.length > 254 || !['viewer','practice_admin'].includes(input?.role)) return reply(400, {error:'Enter a valid practice, email address and role.'});
   const key = env('BREVO_API_KEY'), from = env('INVITE_FROM_EMAIL');
   if (!key || !emailPattern.test(from || '')) return reply(503, {error:'Invitation email is not configured yet. Your administrator needs to connect the email service. No email was sent.'});
   // Limits repeated sends by an administrator. Provider account limits also apply.
   const {count,error:limitError} = await client.from('iq_audit').select('id',{count:'exact',head:true}).eq('actor',identity.user.id).eq('action','invitation_created').gte('occurred_at',new Date(Date.now()-3600000).toISOString());
   if (limitError) return reply(503, {error:'Could not check invitation limits. No email was sent.'});
   if (count >= 20) return reply(429, {error:'Invitation limit reached. Try again in an hour.'});
   const {data:invitation,error} = await client.rpc('iq_create_invite', {p_practice:input.practice,p_email:email,p_role:input.role});
   if (error) return reply(403, {error:'You do not have permission to invite this user to this practice.'});
   const link = new URL(SITE + '/');
   link.searchParams.set('invite',invitation.token);
   const role = input.role === 'practice_admin' ? 'practice administrator' : 'practice viewer';
   let delivery;
   try {
    delivery = await send('https://api.brevo.com/v3/smtp/email', {method:'POST',headers:{'api-key':key,'Content-Type':'application/json','accept':'application/json'},signal:AbortSignal.timeout(15000),body:JSON.stringify({sender:{name:'Darwin IQ',email:from},to:[{email}],subject:'Your invitation to Darwin IQ',textContent:`You have been invited to Darwin IQ as a ${role}.\n\nOpen your invitation:\n${link.href}\n\nCreate an account using ${email}, or sign in if you already have one. After you verify your email and accept the invitation, your practice administrator must approve your access.\n\nThis invitation expires at ${invitation.expires_at}.\n\nIf you were not expecting this invitation, you can ignore this email.`})});
   } catch { return reply(502, {error:'The email service did not confirm submission. Check its delivery log before retrying.'}); }
   const result = await delivery.json().catch(()=>({}));
   if (!delivery.ok || !result.messageId) return reply(502, {error:'The email service did not accept the invitation. Check the verified sender and email-service settings. No successful send was confirmed.'});
   return reply(200, {email_submitted:true,email,expires_at:invitation.expires_at});
  } catch { return reply(500, {error:'Invitation could not be completed. No successful send was confirmed.'}); }
 };
}
