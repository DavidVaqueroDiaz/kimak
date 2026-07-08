// ============================================================
//  Configuración local del Marcador KIMAK
// ------------------------------------------------------------
//  1) Copia este archivo como  config.js  (en la misma carpeta).
//  2) Pon tu URL real de la web app de Apps Script.
//  3) config.js está en .gitignore: NO se sube al repositorio.
//     Al repartir la extensión, el ZIP sí lleva config.js dentro.
// ============================================================

// URL que termina en /exec, la que te da Apps Script al desplegar.
const MK_API_URL = "https://script.google.com/macros/s/XXXXXXXXXXXXXXXXXXXX/exec";

// Nombre que se guarda junto a cada marca (para saber quién marcó).
// Opcional: déjalo "" si no quieres personalizarlo.
const MK_USUARIO = "";

// Token compartido. Solo si lo activaste en Apps Script (const TOKEN).
// Debe coincidir exactamente. Déjalo "" si no usas token.
const MK_TOKEN = "";

// Cada cuántos milisegundos se pregunta a la API por marcas de otros.
// 30000 = 30 s. Subir si queréis menos tráfico.
const MK_INTERVALO_MS = 30000;
