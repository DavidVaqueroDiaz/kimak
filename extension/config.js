// ============================================================
//  Configuración del Marcador KIMAK  (archivo real)
// ------------------------------------------------------------
//  Ya lleva la URL de la web app puesta. Revisa MK_TOKEN abajo.
// ============================================================

// URL de la web app de Apps Script (termina en /exec).
const MK_API_URL = "https://script.google.com/macros/s/AKfycbxqY4H-bTyUHLM2VJOt4mxHT2HaAad3HLUXdpXqTLKBrU7VotEcSfFmPp0q2fBotcQV/exec";

// Nombre que se guarda junto a cada marca (opcional). Pon tu nombre si quieres.
const MK_USUARIO = "";

// Contraseña. Debe ser EXACTAMENTE la misma que la constante TOKEN de Apps Script.
//   - Si en Apps Script dejaste  const TOKEN = '';  (sin contraseña) -> deja esto como "".
//   - Si pusiste una contraseña  ej. const TOKEN = 'kimak-2026';     -> pon aquí "kimak-2026".
const MK_TOKEN = "";

// Cada cuántos milisegundos se pregunta por marcas de otros (30000 = 30 s).
const MK_INTERVALO_MS = 30000;
