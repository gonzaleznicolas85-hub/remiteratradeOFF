// netlify/functions/appscript-proxy.js
// --- VALORES FIJOS (hardcode) ---
// ⚠️ CLON "remiteratradeOFF": pegar acá la URL /exec del NUEVO deployment de Apps Script
// (el que quede colgado del Google Sheet duplicado de esta área), una vez creado.
const APPSCRIPT_URL = "PEGAR_AQUI_URL_EXEC_APPSCRIPT_NUEVO";
const API_TOKEN     = "REMITOSDADIGITAL-OFF";
// --------------------------------

exports.handler = async (event, context) => {
  const corsHeaders = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Headers": "Content-Type, X-Token",
    "Access-Control-Allow-Methods": "POST,GET,OPTIONS"
  };

  if (event.httpMethod === "OPTIONS") {
    return { statusCode: 200, headers: corsHeaders, body: "" };
  }

  /* -------------------------------------------------------
   * GET: lectura de datos (stock, historial)
   * ------------------------------------------------------- */
  if (event.httpMethod === "GET") {
    try {
      const action = (event.queryStringParameters && event.queryStringParameters.action) || "";
      if (!action) {
        return {
          statusCode: 400,
          headers: { ...corsHeaders, "Content-Type": "application/json" },
          body: JSON.stringify({ ok: false, error: "Falta parámetro 'action'" })
        };
      }

      const url = `${APPSCRIPT_URL}?action=${encodeURIComponent(action)}&token=${encodeURIComponent(API_TOKEN)}`;
      const res = await fetch(url, { method: "GET" });
      const text = await res.text();
      let data;
      try { data = JSON.parse(text); } catch (e) { data = { ok: false, error: "Respuesta no JSON del Apps Script", raw: text }; }

      return {
        statusCode: 200,
        headers: { ...corsHeaders, "Content-Type": "application/json" },
        body: JSON.stringify(data)
      };
    } catch (err) {
      return {
        statusCode: 500,
        headers: { ...corsHeaders, "Content-Type": "application/json" },
        body: JSON.stringify({ ok: false, error: String(err && err.stack || err) })
      };
    }
  }

  if (event.httpMethod !== "POST") {
    return { statusCode: 405, headers: { ...corsHeaders, "Content-Type":"application/json" },
             body: JSON.stringify({ ok:false, error:"Método no permitido" }) };
  }

  try {
    const contentType = (event.headers["content-type"] || event.headers["Content-Type"] || "").toLowerCase();
    let payload = { header: {}, lines: [] };

    if (contentType.includes("application/json")) {
      // Reenviamos el body completo tal cual (permite acciones nuevas como
      // uploadImagen/sumarStock sin tener que listar cada campo acá).
      const parsed = JSON.parse(event.body || "{}");
      payload = { ...parsed };
    } else {
      // application/x-www-form-urlencoded (desde URLSearchParams) u otros
      const params = new URLSearchParams(event.body || "");
      const get = (k) => params.get(k);
      if (get("header"))     try { payload.header = JSON.parse(get("header")); } catch {}
      if (get("lines"))      try { payload.lines  = JSON.parse(get("lines"));  } catch {}
      if (get("producto"))   try { payload.producto = JSON.parse(get("producto")); } catch {}
      if (get("action"))     payload.action = get("action");
      if (get("sku"))        payload.sku = get("sku");
      if (get("cantidad"))   payload.cantidad = get("cantidad");
      if (get("imagenBase64")) payload.imagenBase64 = get("imagenBase64");
      if (get("imagenName"))   payload.imagenName = get("imagenName");
    }

    payload.token = API_TOKEN;

    const res = await fetch(APPSCRIPT_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json", "X-Token": API_TOKEN },
      body: JSON.stringify(payload)
    });

    const text = await res.text();
    let data;
    try { data = JSON.parse(text); } catch (e) { data = { ok:false, error:"Respuesta no JSON del Apps Script", raw:text }; }

    return {
      statusCode: 200,
      headers: { ...corsHeaders, "Content-Type":"application/json" },
      body: JSON.stringify(data)
    };
  } catch (err) {
    return {
      statusCode: 500,
      headers: { ...corsHeaders, "Content-Type":"application/json" },
      body: JSON.stringify({ ok:false, error: String(err && err.stack || err) })
    };
  }
};
