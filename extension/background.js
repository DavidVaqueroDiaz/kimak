// ============================================================
//  Marcador KIMAK — Service worker (fondo)
// ------------------------------------------------------------
//  Hace las llamadas HTTPS a la API de Apps Script.
//
//  ¿Por qué aquí y no en content.js? En Manifest V3, un fetch
//  lanzado desde el content script usa el origen de la página
//  (sig-design.kimak.com) y choca con CORS. Desde el service
//  worker, con "host_permissions" concedidos, el fetch va con el
//  origen de la extensión y no hay problema de CORS.
//
//  El content.js le manda mensajes; aquí se resuelven.
// ============================================================

/** GET: pide la lista de IDs marcados. */
async function apiObtener(url, token) {
  const u = token
    ? url + (url.includes('?') ? '&' : '?') + 'token=' + encodeURIComponent(token)
    : url;
  const res = await fetch(u, { method: 'GET' });
  return res.json();
}

/** POST: marca o desmarca un ID. text/plain para evitar el preflight CORS. */
async function apiAlternar(url, payload) {
  const res = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify(payload),
  });
  return res.json();
}

/** Enruta los mensajes del content script. */
async function manejar(msg) {
  if (!msg || !msg.url || String(msg.url).indexOf('XXXX') !== -1) {
    return { ok: false, error: 'sin_url' };
  }

  if (msg.tipo === 'obtener') {
    return apiObtener(msg.url, msg.token);
  }

  if (msg.tipo === 'alternar') {
    return apiAlternar(msg.url, {
      id: msg.id,
      marcado: msg.marcado,
      usuario: msg.usuario,
      token: msg.token,
    });
  }

  return { ok: false, error: 'tipo_desconocido' };
}

chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
  manejar(msg)
    .then(sendResponse)
    .catch((e) => sendResponse({ ok: false, error: String(e) }));
  return true; // respuesta asíncrona
});
