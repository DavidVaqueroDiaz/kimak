// ============================================================
//  Marcador KIMAK — Otro departamento (compartido)
//  Marca tareas que hace otro departamento. Ahora las marcas se
//  guardan en una hoja compartida (vía Apps Script), así lo que
//  marca uno lo ven todos.
//
//  Estrategia de robustez (para no saturar la API ni el SIG):
//   - Caché en memoria (Set) como fuente de verdad para pintar.
//   - El MutationObserver (scroll infinito, filtros) SOLO repinta
//     desde la caché: no llama a la red.
//   - La red se toca en momentos concretos: al cargar, cada X seg,
//     y al marcar/desmarcar (optimista, con reversión si falla).
//   - Fallback a chrome.storage.local si la API no responde.
// ============================================================

const CLAVE_LOCAL = "mk_tareas_otro_depto"; // copia de seguridad local

// Config viene de config.js (cargado antes que este archivo).
const API_URL  = (typeof MK_API_URL !== "undefined") ? MK_API_URL : "";
const USUARIO  = (typeof MK_USUARIO !== "undefined" && MK_USUARIO) ? MK_USUARIO : "anon";
const TOKEN    = (typeof MK_TOKEN !== "undefined") ? MK_TOKEN : "";
const INTERVALO = (typeof MK_INTERVALO_MS !== "undefined") ? MK_INTERVALO_MS : 30000;

const apiConfigurada = API_URL && API_URL.indexOf("XXXX") === -1;

// Fuente de verdad en memoria: IDs marcados.
let cache = new Set();

// ---------- Copia de seguridad local (fallback) ----------

function leerLocal() {
  return new Promise((resolve) => {
    chrome.storage.local.get([CLAVE_LOCAL], (r) => {
      resolve(new Set((r[CLAVE_LOCAL] || []).map(String)));
    });
  });
}
function guardarLocal(set) {
  chrome.storage.local.set({ [CLAVE_LOCAL]: Array.from(set) });
}

// ---------- Puente con el service worker (red) ----------

function pedirAlFondo(msg) {
  return new Promise((resolve) => {
    try {
      chrome.runtime.sendMessage(msg, (resp) => {
        if (chrome.runtime.lastError) {
          resolve({ ok: false, error: chrome.runtime.lastError.message });
        } else {
          resolve(resp || { ok: false, error: "sin_respuesta" });
        }
      });
    } catch (e) {
      resolve({ ok: false, error: String(e) });
    }
  });
}

function apiObtener() {
  return pedirAlFondo({ tipo: "obtener", url: API_URL, token: TOKEN });
}
function apiAlternar(id, marcado) {
  return pedirAlFondo({
    tipo: "alternar", url: API_URL, token: TOKEN,
    id: id, marcado: marcado, usuario: USUARIO,
  });
}

// ---------- Helpers de DOM (sin cambios de lógica) ----------

// Extrae el id único de la fila a partir de id="listItem_XXXX".
function idDeFila(tr) {
  const m = (tr.id || "").match(/listItem_(\d+)/);
  return m ? m[1] : null;
}

// Localiza el recuadro del número SIG dentro de la fila.
function recuadroSig(tr) {
  const baloon = tr.querySelector('a.baloon[onclick*="marcaSelecionado"]');
  if (baloon) return baloon;
  const candidatos = tr.querySelectorAll("a, span, div, button");
  for (const el of candidatos) {
    const t = (el.textContent || "").trim();
    if (/^\d{5,}\s*\|\s*\d{3,}$/.test(t)) return el;
  }
  return null;
}

// Aplica o quita el aspecto visual de "marcada".
function pintar(tr, marcada) {
  const sig = recuadroSig(tr);
  if (marcada) {
    tr.classList.add("mk-fila-marcada");
    if (sig) sig.classList.add("mk-otro-depto");
  } else {
    tr.classList.remove("mk-fila-marcada");
    if (sig) sig.classList.remove("mk-otro-depto");
  }
  const btn = tr.querySelector(".mk-btn");
  if (btn) {
    btn.classList.toggle("mk-activo", marcada);
    btn.title = marcada
      ? "Marcada: otro departamento (clic para quitar)"
      : "Marcar: otro departamento";
  }
}

// ---------- Marcar / desmarcar (optimista + red) ----------

async function alternar(tr) {
  const id = idDeFila(tr);
  if (!id) return;

  const marcar = !cache.has(id);

  // 1) Optimista: actualizamos caché, UI y copia local ya.
  if (marcar) cache.add(id); else cache.delete(id);
  pintarPorId(id, marcar);
  guardarLocal(cache);

  if (!apiConfigurada) {
    // Sin API configurada: funciona en modo local, como antes.
    return;
  }

  // 2) Confirmamos con el servidor.
  const r = await apiAlternar(id, marcar);
  if (!r || !r.ok) {
    // Revertimos si el servidor no aceptó.
    if (marcar) cache.delete(id); else cache.add(id);
    pintarPorId(id, !marcar);
    guardarLocal(cache);
    avisar("No se pudo guardar en el servidor. Reintenta.");
  }
}

// ---------- Preparación de filas ----------

function prepararFila(tr) {
  const id = idDeFila(tr);
  if (!id) return;

  const sig = recuadroSig(tr);
  const chkBlanco = tr.querySelector('input[type="checkbox"]');
  const contenedorChk = chkBlanco ? chkBlanco.closest('.custom-controls-stacked') : null;

  // Cuadradito naranja a la izquierda del checkbox blanco.
  if ((sig || chkBlanco) && !tr.querySelector(".mk-btn")) {
    const btn = document.createElement("span");
    btn.className = "mk-btn";
    btn.title = "Marcar: otro departamento";
    btn.addEventListener("click", (e) => {
      e.stopPropagation();
      e.preventDefault();
      alternar(tr);
    });
    if (contenedorChk) {
      contenedorChk.insertAdjacentElement("beforebegin", btn);
    } else if (chkBlanco) {
      chkBlanco.insertAdjacentElement("beforebegin", btn);
    } else if (sig) {
      sig.insertAdjacentElement("afterend", btn);
    }
  }

  // Clic derecho sobre el recuadro SIG = marcar/desmarcar.
  if (sig && !sig.dataset.mkListo) {
    sig.dataset.mkListo = "1";
    sig.addEventListener("contextmenu", (e) => {
      e.preventDefault();
      e.stopPropagation();
      alternar(tr);
    });
  }

  pintar(tr, cache.has(id));
}

// Repinta una fila concreta buscándola por id (tras un toggle).
function pintarPorId(id, marcada) {
  const tr = document.getElementById("listItem_" + id);
  if (tr) pintar(tr, marcada);
}

// Recorre las filas visibles: prepara botones y pinta desde la caché.
function procesarTabla() {
  const filas = document.querySelectorAll('tr[id^="listItem_"]');
  filas.forEach((tr) => prepararFila(tr));
}

// Repinta todo desde la caché (tras sincronizar con el servidor).
function repintarTodo() {
  const filas = document.querySelectorAll('tr[id^="listItem_"]');
  filas.forEach((tr) => {
    const id = idDeFila(tr);
    if (id) pintar(tr, cache.has(id));
  });
}

// ---------- Sincronización con el servidor ----------

async function sincronizar() {
  if (!apiConfigurada) {
    cache = await leerLocal();
    repintarTodo();
    return;
  }
  const r = await apiObtener();
  if (r && r.ok && Array.isArray(r.marcadas)) {
    cache = new Set(r.marcadas.map(String));
    guardarLocal(cache); // snapshot para el fallback
  } else {
    // API caída: usamos la última copia local conocida.
    cache = await leerLocal();
  }
  repintarTodo();
}

// ---------- Aviso discreto ----------

let avisoTimer = null;
function avisar(texto) {
  let caja = document.getElementById("mk-aviso");
  if (!caja) {
    caja = document.createElement("div");
    caja.id = "mk-aviso";
    document.body.appendChild(caja);
  }
  caja.textContent = texto;
  caja.classList.add("mk-aviso-visible");
  clearTimeout(avisoTimer);
  avisoTimer = setTimeout(() => caja.classList.remove("mk-aviso-visible"), 4000);
}

// ---------- Arranque ----------

// Observa cambios de la tabla (filtrar, paginar, scroll). Solo repinta
// desde la caché; NO llama a la red.
const obs = new MutationObserver(() => {
  clearTimeout(window.__mkTimer);
  window.__mkTimer = setTimeout(procesarTabla, 300);
});
obs.observe(document.body, { childList: true, subtree: true });

// Primera pasada: pinta con lo local al instante y luego sincroniza.
(async function inicio() {
  cache = await leerLocal();
  procesarTabla();
  await sincronizar();
  if (apiConfigurada && INTERVALO > 0) {
    setInterval(sincronizar, INTERVALO);
  }
})();
