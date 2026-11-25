const crypto=require('node:crypto');
module.exports=async function(context,req){
const out=(status,body)=>{context.res={status,headers:{'Content-Type':'application/json','Cache-Control':'no-store'},body}};
try{
const username=process.env.EXENTIA_AUTH_USERNAME||'';
const parts=(process.env.EXENTIA_AUTH_PASSWORD_HASH||'').split('$');
const key=Buffer.from(process.env.EXENTIA_AUTH_SESSION_KEY||'','base64');
const configured=!!username&&parts.length===4&&parts[0]==='pbkdf2-sha256'&&parts[1]==='600000'&&Buffer.from(parts[2]||'','base64').length===32&&Buffer.from(parts[3]||'','base64').length===32&&key.length===32;
const token=await fetch('https://login.microsoftonline.com/'+encodeURIComponent(process.env.EXENTIA_TENANT_ID)+'/oauth2/v2.0/token',{method:'POST',body:new URLSearchParams({client_id:process.env.EXENTIA_CLIENT_ID,client_secret:process.env.EXENTIA_CLIENT_SECRET,scope:'https://graph.microsoft.com/.default',grant_type:'client_credentials'})});
if(!token.ok)return out(503,{configured,graph:false});
const access=(await token.json()).access_token;
const site=encodeURIComponent(process.env.EXENTIA_SITE_ID).replace(/%2C/g,',');
const r=await fetch('https://graph.microsoft.com/v1.0/sites/'+site+'/lists?$select=id&$top=1',{headers:{Authorization:'Bearer '+access}});
const result={configured,graph:true,listsReadable:r.ok,listsStatus:r.status};
if(req.query?.write!=='1')return out(200,result);
const authSite='exentiaperu.sharepoint.com,83fe5fa6-0387-4129-aa04-df303819ba9f,b58b1829-7753-4222-adb9-f8688809869b';
const itemUrl='https://graph.microsoft.com/v1.0/sites/'+authSite+'/lists/17f3bba3-a178-4d54-996b-fce9c541e482/items/1';
const headers={Authorization:'Bearer '+access,'Content-Type':'application/json'};
const read=await fetch(itemUrl+'?$expand=fields',{headers});
result.authReadStatus=read.status;
if(!read.ok)return out(200,result);
const item=await read.json();
const etag=item.eTag||item['@odata.etag'];
const state=item.fields?.AuthState;
if(item.fields?.Title!=='admin'||typeof state!=='string'||!etag)return out(200,{...result,recordValid:false});
try{const parsed=JSON.parse(state);result.recordValid=parsed.version===1&&Array.isArray(parsed.sessions)&&Number.isInteger(parsed.failedAttempts)&&Number.isFinite(parsed.lockedUntil);}catch(e){result.recordValid=false;}
if(!result.recordValid)return out(200,result);
const updated=await fetch(itemUrl+'/fields',{method:'PATCH',headers:{...headers,'If-Match':etag},body:JSON.stringify({AuthState:state})});
result.authWriteStatus=updated.status;
if(updated.ok){
 const after=await fetch(itemUrl+'?$expand=fields',{headers});
 if(after.ok){const data=await after.json();const newTag=data.eTag||data['@odata.etag'];result.statePreserved=data.fields?.AuthState===state;result.etagChanged=!!newTag&&newTag!==etag;
 if(result.etagChanged){const stale=await fetch(itemUrl+'/fields',{method:'PATCH',headers:{...headers,'If-Match':etag},body:JSON.stringify({AuthState:state})});result.staleUpdateRejected=stale.status===412;}
 }
}
return out(200,result);
}catch(e){return out(503,{ok:false,message:'No se pudo verificar la infraestructura.'})}
};