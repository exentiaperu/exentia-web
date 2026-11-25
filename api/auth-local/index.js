const crypto=require('node:crypto');
const {promisify}=require('node:util');
const derive=promisify(crypto.pbkdf2);
const origin='https://exentia.pe';
const cookie='__Host-exentia_session';
const ttl=8*60*60*1000;
const url='https://graph.microsoft.com/v1.0/sites/exentiaperu.sharepoint.com,83fe5fa6-0387-4129-aa04-df303819ba9f,b58b1829-7753-4222-adb9-f8688809869b/lists/17f3bba3-a178-4d54-996b-fce9c541e482/items/1';
function config(){
 const p=(process.env.EXENTIA_AUTH_PASSWORD_HASH||'').split('$');
 const key=Buffer.from(process.env.EXENTIA_AUTH_SESSION_KEY||'','base64');
 const salt=Buffer.from(p[2]||'','base64'), hash=Buffer.from(p[3]||'','base64');
 const user=process.env.EXENTIA_AUTH_USERNAME;
 if(!user||p.length!==4||p[0]!=='pbkdf2-sha256'||p[1]!=='600000'||salt.length!==32||hash.length!==32||key.length!==32)throw Error('config');
 return {user,key,salt,hash};
}
async function headers(){
 const r=await fetch('https://login.microsoftonline.com/'+encodeURIComponent(process.env.EXENTIA_TENANT_ID)+'/oauth2/v2.0/token',{method:'POST',body:new URLSearchParams({client_id:process.env.EXENTIA_CLIENT_ID,client_secret:process.env.EXENTIA_CLIENT_SECRET,scope:'https://graph.microsoft.com/.default',grant_type:'client_credentials'}),signal:AbortSignal.timeout(10000)});
 if(!r.ok)throw Error('token');
 return {Authorization:'Bearer '+(await r.json()).access_token,'Content-Type':'application/json'};
}
async function read(h){
 const r=await fetch(url+'?$expand=fields',{headers:h,signal:AbortSignal.timeout(10000)});
 if(!r.ok)throw Error('read');
 const d=await r.json(), etag=d.eTag||d['@odata.etag'], s=JSON.parse(d.fields.AuthState);
 if(!etag||d.fields.Title!=='admin'||s.version!==1||!Number.isSafeInteger(s.failedAttempts)||s.failedAttempts<0||!Number.isSafeInteger(s.lockedUntil)||!Array.isArray(s.sessions)||s.sessions.some(x=>!x||typeof x.id!=='string'||!Number.isSafeInteger(x.exp)))throw Error('state');
 s.sessions=s.sessions.filter(x=>x.exp>Date.now());
 return {s,etag};
}
async function change(h,fn){
 for(let i=0;i<5;i++){
 const {s,etag}=await read(h); const result=fn(s);
 if(result.skip)return result;
 const r=await fetch(url+'/fields',{method:'PATCH',headers:{...h,'If-Match':etag},body:JSON.stringify({AuthState:JSON.stringify(s)}),signal:AbortSignal.timeout(10000)});
 if(r.ok)return result;
 if(r.status!==412)throw Error('write');
 }
 throw Error('conflict');
}
function digest(value,key){return crypto.createHmac('sha256',key).update(value).digest('hex');}
function token(req){const m=(req.headers.cookie||'').match(/(?:^|;\s*)__Host-exentia_session=([A-Za-z0-9_-]{43})(?:;|$)/);return m?.[1]||'';}
function reply(context,status,body,setCookie){context.res={status,headers:{'Content-Type':'application/json; charset=utf-8','Cache-Control':'no-store','X-Content-Type-Options':'nosniff',...(setCookie?{'Set-Cookie':setCookie}:{})},body};}
function set(value,age){return cookie+'='+value+'; Path=/; Secure; HttpOnly; SameSite=Strict; Max-Age='+age;}
async function handler(context,req){
 try{
 const c=config(), method=req.method.toUpperCase();
 if(method==='GET'){
 const t=token(req); if(!t)return reply(context,200,{authenticated:false});
 const {s}=await read(await headers());
 return reply(context,200,{authenticated:s.sessions.some(x=>x.id===digest(t,c.key))});
 }
 if(method!=='POST'||req.headers.origin!==origin||!(req.headers['content-type']||'').toLowerCase().startsWith('application/json'))return reply(context,403,{ok:false});
 const b=req.body;
 if(!b||typeof b!=='object')return reply(context,400,{ok:false});
 if(b.action==='logout'){
 const t=token(req);
 if(t){const h=await headers();await change(h,s=>{s.sessions=s.sessions.filter(x=>x.id!==digest(t,c.key));return {};});}
 return reply(context,200,{ok:true},set('',0));
 }
 if(b.action!=='login'||typeof b.username!=='string'||typeof b.password!=='string'||b.username.length>254||b.password.length>1024)return reply(context,400,{ok:false});
 const h=await headers();
 // Reserve each attempt atomically before password verification, preventing concurrent guesses from bypassing the limit.
 const reserved=await change(h,s=>{
 const now=Date.now();
 if(s.lockedUntil>now)return {skip:true,locked:true};
 if(s.lockedUntil!==0){s.failedAttempts=0;s.lockedUntil=0;}
 s.failedAttempts++;
 if(s.failedAttempts>=5)s.lockedUntil=now+15*60*1000;
 return {};
 });
 if(reserved.locked)return reply(context,429,{ok:false,message:'Acceso temporalmente bloqueado. Intenta en 15 minutos.'});
 const actual=await derive(b.password,c.salt,600000,32,'sha256');
 const valid=crypto.timingSafeEqual(actual,c.hash)&&b.username.trim().toLowerCase()===c.user.toLowerCase();
 if(!valid)return reply(context,401,{ok:false,message:'Credenciales incorrectas.'});
 const t=crypto.randomBytes(32).toString('base64url');
 await change(h,s=>{s.failedAttempts=0;s.lockedUntil=0;s.sessions=s.sessions.slice(-19);s.sessions.push({id:digest(t,c.key),exp:Date.now()+ttl});return {};});
 return reply(context,200,{ok:true},set(t,ttl/1000));
 }catch(e){return reply(context,503,{ok:false,message:'Servicio temporalmente no disponible.'});}
}
module.exports=handler;