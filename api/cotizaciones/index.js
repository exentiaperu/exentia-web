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

function graphFileUrl() {
  const siteId = process.env.EXENTIA_SITE_ID;
  if (!siteId) throw new Error("EXENTIA_SITE_ID no configurado");
  const encoded = encodeURIComponent(siteId).replace(/%2C/g, ",");
  return `https://graph.microsoft.com/v1.0/sites/${encoded}/drive/root:/EXENTIA/Cotizaciones/cotizaciones.json:/content`;
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
      context.res = json(200, { ok:true });
      return;
    }

    context.res = json(405, { ok:false, error:"Método no permitido" });
  } catch (e) {
    context.log.error(e);
    context.res = json(500, { ok:false, error:e.message || "Error interno" });
  }
};