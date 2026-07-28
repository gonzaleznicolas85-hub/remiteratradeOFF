Remitos – Netlify Functions v1 (Node) – FIX para "Failed to fetch"

Qué cambia:
- Función reescrita con `exports.handler` (Node/Functions v1), compatible con la mayoría de sites.
- Convierte cualquier request (JSON o x-www-form-urlencoded) a JSON para Apps Script.

Pasos:
1) Subí estos archivos (o reemplazá tu función) y hacé deploy.
2) Tu front sigue posteando a: /.netlify/functions/appscript-proxy
3) No necesitás env vars: la URL y token están hardcodeados del lado servidor.

Test en consola:
const form = new URLSearchParams();
form.append("header", JSON.stringify({ fecha: new Date().toISOString().slice(0,10), punto_venta:"MDP-01", usuario:"nico", obs:"prueba" }));
form.append("lines", JSON.stringify([{ sku:"10331", descripcion:"Carne", tipo:"empanada", marca:"La Vaska", cantidad:2 }]));
fetch("/.netlify/functions/appscript-proxy", { method:"POST", body: form }).then(r=>r.json()).then(console.log);
