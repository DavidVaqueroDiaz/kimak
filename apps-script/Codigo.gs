/**
 * ============================================================
 *  Marcador KIMAK — API web (Google Apps Script)
 * ============================================================
 *
 *  Base de datos = una hoja de Google Sheets (una fila por marca).
 *  Esta API la consulta la extensión del navegador por HTTPS:
 *
 *    GET  /exec            -> devuelve la lista de IDs marcados
 *    POST /exec  {id,...}  -> marca o desmarca un ID
 *
 *  Puntos clave resueltos aquí:
 *   - Concurrencia: LockService serializa las escrituras, así dos
 *     personas marcando a la vez no se pisan.
 *   - CORS: la extensión envía el POST como text/plain para evitar
 *     el preflight OPTIONS que Apps Script no maneja.
 *   - El dato (número SIG) no es sensible; aun así hay un TOKEN
 *     opcional para que solo escriba quien tenga la clave.
 *
 *  Cómo desplegar: ver README.md (Implementar > Nueva implementación
 *  > Aplicación web).
 * ============================================================
 */

// --- Configuración ---------------------------------------------------------

// Nombre de la pestaña donde se guardan las marcas. Se crea sola si no existe.
const NOMBRE_HOJA = 'Marcas';

// Token compartido opcional. Si lo dejas vacío ("") no se comprueba nada.
// Si pones un valor, la extensión debe enviar el mismo token (config.js).
const TOKEN = '';

// --- Utilidades ------------------------------------------------------------

/** Devuelve la hoja de marcas, creándola con cabecera si hace falta. */
function hoja_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let h = ss.getSheetByName(NOMBRE_HOJA);
  if (!h) {
    h = ss.insertSheet(NOMBRE_HOJA);
    h.appendRow(['id', 'marcado', 'usuario', 'actualizado']);
    h.setFrozenRows(1);
  }
  return h;
}

/** Respuesta JSON estándar. */
function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Comprueba el token si está configurado. Devuelve true si pasa. */
function tokenOk_(valor) {
  if (!TOKEN) return true;
  return String(valor || '') === TOKEN;
}

// --- Endpoints -------------------------------------------------------------

/**
 * GET: devuelve los IDs actualmente marcados.
 * Ábrelo en el navegador para probar que la web app responde.
 */
function doGet(e) {
  try {
    const params = (e && e.parameter) || {};
    if (!tokenOk_(params.token)) {
      return json_({ ok: false, error: 'token' });
    }

    const filas = hoja_().getDataRange().getValues(); // incluye cabecera
    const marcadas = [];
    for (let i = 1; i < filas.length; i++) {
      const id = filas[i][0];
      const marcado = filas[i][1];
      if (id !== '' && marcado === true) marcadas.push(String(id));
    }
    return json_({ ok: true, marcadas: marcadas, total: marcadas.length });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/**
 * POST: marca o desmarca un ID.
 * Cuerpo esperado (text/plain con JSON dentro):
 *   { "id": "12345", "marcado": true, "usuario": "nombre", "token": "..." }
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000); // espera hasta 20 s a que se libere
  } catch (err) {
    return json_({ ok: false, error: 'ocupado, reintenta' });
  }

  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');

    if (!tokenOk_(body.token)) {
      return json_({ ok: false, error: 'token' });
    }

    const id = String(body.id || '').trim();
    if (!id) return json_({ ok: false, error: 'falta id' });

    const marcado = body.marcado === true;
    const usuario = String(body.usuario || 'anon').slice(0, 60);
    const ahora = new Date();

    const h = hoja_();
    const filas = h.getDataRange().getValues();
    let fila = -1;
    for (let i = 1; i < filas.length; i++) {
      if (String(filas[i][0]) === id) { fila = i + 1; break; } // +1: hoja es 1-based
    }

    if (fila === -1) {
      h.appendRow([id, marcado, usuario, ahora]);
    } else {
      // Actualiza marcado / usuario / actualizado, deja el id como está.
      h.getRange(fila, 2, 1, 3).setValues([[marcado, usuario, ahora]]);
    }

    return json_({ ok: true, id: id, marcado: marcado });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
