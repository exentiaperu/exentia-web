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
const base='https://graph.microsoft.com/v1.0/sites/'+site+'/lists';
const headers={Authorization:'Bearer '+access,'Content-Type':'application/json'};
const name='EXENTIA Auth Probe '+crypto.randomUUID();
const created=await fetch(base,{method:'POST',headers,body:JSON.stringify({displayName:name,list:{template:'genericList'}})});
result.listCreateStatus=created.status;
if(!created.ok)return out(200,result);
const id=(await created.json()).id;
if(!id)return out(503,{...result,cleanup:false});
const url=base+'/'+encodeURIComponent(id);
try{
 const item=await fetch(url+'/items',{method:'POST',headers,body:JSON.stringify({fields:{Title:'Auth infrastructure probe'}})});
 result.itemCreateStatus=item.status;
 if(item.ok){
 const obj=await item.json();
 const itemUrl=url+'/items/'+encodeURIComponent(obj.id);
 const read=await fetch(itemUrl+'?$expand=fields',{headers});
 result.itemReadStatus=read.status;
 if(read.ok){
 const data=await read.json();
 const etag=data.eTag||data['@odata.etag'];
 if(etag){
 const updated=await fetch(itemUrl+'/fields',{method:'PATCH',headers:{...headers,'If-Match':etag},body:JSON.stringify({Title:'Auth infrastructure probe updated'})});
 result.conditionalUpdateStatus=updated.status;
 if(updated.ok){
 const stale=await fetch(itemUrl+'/fields',{method:'PATCH',headers:{...headers,'If-Match':etag},body:JSON.stringify({Title:'Stale update test'})});
 result.staleUpdateRejected=stale.status===412;
 }
 }else result.etagPresent=false;
 }
 }
}finally{
 const cleanup=await fetch(url,{method:'DELETE',headers});
 result.cleanupStatus=cleanup.status;result.cleanup=cleanup.ok;
}
return out(200,result);
}catch(e){return out(503,{ok:false,message:'No se pudo verificar la infraestructura.'})}
};