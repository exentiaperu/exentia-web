const { PDFDocument, StandardFonts, rgb } = require("pdf-lib");
async function getToken() {
  const tenant = process.env.EXENTIA_TENANT_ID;
  const clientId = process.env.EXENTIA_CLIENT_ID;
  const clientSecret = process.env.EXENTIA_CLIENT_SECRET;
  if (!tenant || !clientId || !clientSecret) throw new Error("Configuración incompleta de Microsoft Graph");

  const body = new URLSearchParams({
    client_id: clientId,
    client_secret: clientSecret,
    scope: "https://graph.microsoft.com/.default",
    grant_type: "client_credentials"
  });

  const r = await fetch(`https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body
  });
  if (!r.ok) throw new Error("No se pudo obtener token de Microsoft Graph");
  return (await r.json()).access_token;
}

function graphBase() {
  const siteId = process.env.EXENTIA_SITE_ID;
  if (!siteId) throw new Error("EXENTIA_SITE_ID no configurado");
  const encoded = encodeURIComponent(siteId).replace(/%2C/g, ",");
  return `https://graph.microsoft.com/v1.0/sites/${encoded}/drive/root:/EXENTIA/Cotizaciones`;
}
function graphFileUrl() {
  return graphBase() + "/cotizaciones.json:/content";
}
function graphNamedFileUrl(name) {
  return graphBase() + "/" + encodeURIComponent(name) + ":/content";
}

async function readQuotes(token) {
  const r = await fetch(graphFileUrl(), { headers: { Authorization: `Bearer ${token}` } });
  if (r.status === 404) return [];
  if (!r.ok) throw new Error(`No se pudo leer cotizaciones (${r.status})`);
  const text = await r.text();
  if (!text.trim()) return [];
  const data = JSON.parse(text);
  return Array.isArray(data) ? data : [];
}

async function writeQuotes(token, quotes) {
  const r = await fetch(graphFileUrl(), {
    method: "PUT",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify(quotes, null, 2)
  });
  if (!r.ok) throw new Error(`No se pudo guardar cotizaciones (${r.status})`);
}

function safeName(v) {
  return String(v || "CLIENTE")
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-zA-Z0-9-_ ]/g, "")
    .trim().replace(/\s+/g, "-").slice(0, 50) || "CLIENTE";
}
function money(v){ return "S/ " + Number(v || 0).toLocaleString("es-PE", {minimumFractionDigits:2, maximumFractionDigits:2}); }
function wrap(text, max=80){
  const words=String(text||"").split(/\s+/), lines=[]; let line="";
  for(const w of words){ const test=(line+" "+w).trim(); if(test.length>max && line){lines.push(line);line=w}else line=test; }
  if(line)lines.push(line); return lines;
}
async function buildQuotePdf(quote) {
  const pdf = await PDFDocument.create();
  let page = pdf.addPage([595.28, 841.89]);
  const regular = await pdf.embedFont(StandardFonts.Helvetica);
  const bold = await pdf.embedFont(StandardFonts.HelveticaBold);
  let y = 790;
  const draw=(text,size=10,isBold=false,indent=0)=>{
    if(y<70){page=pdf.addPage([595.28,841.89]);y=790;}
    page.drawText(String(text||""),{x:50+indent,y,size,font:isBold?bold:regular,color:rgb(0.05,0.12,0.23)});
    y-=size+7;
  };
  draw("EXENTIA",20,true); draw("COTIZACION",14,true); y-=8;
  draw("Cliente: " + (quote.empresa||"-"),11,true);
  draw("RUC: " + (quote.ruc||"-"));
  draw("Contacto: " + (quote.contacto||"-"));
  draw("Servicio: " + (quote.servicio||"-"));
  draw("Version: " + (quote.version||1));
  draw("Estado: " + (quote.estado||"Borrador"));
  draw("Fecha: " + new Date(quote.updatedAt||quote.createdAt||Date.now()).toLocaleDateString("es-PE"));
  y-=10;
  draw("PRECIO OFRECIDO: " + money(quote.precioOfrecido||quote.recomendado) + " + IGV",14,true);
  y-=10;
  draw("ALCANCE / RESUMEN",11,true);
  for(const line of wrap(quote.resumen||quote.observacionesComerciales||"Sin observaciones",88)) draw(line,9,false,0);
  y-=8;
  draw("Documento generado automaticamente por el panel interno de EXENTIA.",8);
  return await pdf.save();
}
async function uploadQuotePdf(token, quote) {
  const shortId=String(quote.id||"").replace(/-/g,"").slice(0,8);
  const filename=`COT-${shortId}-${safeName(quote.empresa||quote.ruc)}-v${String(quote.version||1).padStart(2,"0")}.pdf`;
  const bytes=await buildQuotePdf(quote);
  const r=await fetch(graphNamedFileUrl(filename),{
    method:"PUT",
    headers:{Authorization:`Bearer ${token}`,"Content-Type":"application/pdf"},
    body:Buffer.from(bytes)
  });
  if(!r.ok) throw new Error(`La cotizacion se guardo, pero no se pudo crear su PDF (${r.status})`);
  return filename;
}
async function deleteQuotePdfs(token,id){
  const shortId=String(id||"").replace(/-/g,"").slice(0,8);
  const r=await fetch(graphBase()+":/children",{headers:{Authorization:`Bearer ${token}`}});
  if(!r.ok)return;
  const data=await r.json();
  const files=(data.value||[]).filter(x=>x.name&&x.name.startsWith("COT-"+shortId+"-")&&x.id);
  for(const file of files){
    await fetch(`https://graph.microsoft.com/v1.0/sites/${encodeURIComponent(process.env.EXENTIA_SITE_ID).replace(/%2C/g,",")}/drive/items/${file.id}`,{
      method:"DELETE",headers:{Authorization:`Bearer ${token}`}
    });
  }
}

function json(status, body) {
  return {
    status,
    headers: { "Content-Type": "application/json; charset=utf-8", "Cache-Control": "no-store" },
    body
  };
}

module.exports = async function (context, req) {
  try {
    const token = await getToken();
    const method = (req.method || "GET").toUpperCase();
    let quotes = await readQuotes(token);

    if (method === "GET") {
      quotes.sort((a,b) => String(b.updatedAt || b.createdAt || "").localeCompare(String(a.updatedAt || a.createdAt || "")));
      context.res = json(200, { ok: true, quotes });
      return;
    }

    const body = req.body || {};
    const now = new Date().toISOString();

    if (method === "POST") {
      const quote = {
        ...body,
        id: body.id || (globalThis.crypto?.randomUUID?.() || (Date.now().toString(36) + Math.random().toString(36).slice(2))),
        version: 1,
        createdAt: now,
        updatedAt: now
      };
      quotes.push(quote);
      await writeQuotes(token, quotes);
      quote.pdfFile = await uploadQuotePdf(token, quote);
      await writeQuotes(token, quotes);
      context.res = json(201, { ok: true, quote });
      return;
    }

    if (method === "PUT") {
      if (!body.id) { context.res = json(400, { ok:false, error:"Falta id" }); return; }
      const i = quotes.findIndex(q => q.id === body.id);
      if (i < 0) { context.res = json(404, { ok:false, error:"Cotización no encontrada" }); return; }
      const previous = quotes[i];
      const quote = {
        ...previous,
        ...body,
        id: previous.id,
        version: Number(previous.version || 1) + 1,
        createdAt: previous.createdAt || now,
        updatedAt: now
      };
      quotes[i] = quote;
      await writeQuotes(token, quotes);
      quote.pdfFile = await uploadQuotePdf(token, quote);
      await writeQuotes(token, quotes);
      context.res = json(200, { ok:true, quote });
      return;
    }

    if (method === "DELETE") {
      const id = body.id || req.query?.id;
      if (!id) { context.res = json(400, { ok:false, error:"Falta id" }); return; }
      const before = quotes.length;
      quotes = quotes.filter(q => q.id !== id);
      if (quotes.length === before) { context.res = json(404, { ok:false, error:"Cotización no encontrada" }); return; }
      await writeQuotes(token, quotes);
      await deleteQuotePdfs(token, id);
      context.res = json(200, { ok:true });
      return;
    }

    context.res = json(405, { ok:false, error:"Método no permitido" });
  } catch (e) {
    context.log.error(e);
    context.res = json(500, { ok:false, error:e.message || "Error interno" });
  }
};