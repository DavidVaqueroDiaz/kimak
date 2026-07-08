# Marcador KIMAK — Otro departamento (compartido)

Extensión de navegador para marcar en el SIG (`sig-design.kimak.com`) qué tareas
son de otro departamento y no revisarlas dos veces. **Lo que marca uno lo ven
todos** los del equipo (5-6 personas).

## Cómo funciona (resumen)

Una extensión de navegador vive en una caja de arena: **solo puede hablar por
HTTPS**, no puede leer/escribir archivos en disco ni en rutas de red
(`\\servidor10\...`). Por eso la base de datos compartida no puede ser un
Excel/JSON en carpeta compartida.

La solución elegida:

```
  Extensión (cada PC)  ──HTTPS──►  Apps Script (web app)  ──►  Google Sheet
       content.js                     doGet / doPost              1 fila = 1 marca
       background.js
```

- **Google Sheet** = base de datos (una fila por marca).
- **Apps Script** asociado = API web (`doGet` / `doPost`) que la extensión
  consulta por HTTPS.
- **Extensión** = pinta el SIG y sincroniza con esa API.

Ventajas resueltas de fábrica:

- **Concurrencia**: `LockService` serializa las escrituras; dos personas
  marcando a la vez no se pisan.
- **CORS**: el POST se envía como `text/plain` para evitar el preflight
  `OPTIONS` que Apps Script no maneja bien. Además el `fetch` se hace desde el
  service worker (origen de la extensión), no desde la página.
- **No hace falta PC encendido**: el código vive en la nube de Google.

## Estructura del repositorio

```
apps-script/
  Codigo.gs           API web (doGet/doPost) + LockService
  appsscript.json     Config del proyecto (zona horaria, tipo de acceso)
extension/
  manifest.json       Manifest V3
  background.js       Service worker: hace los fetch a la API (evita CORS)
  content.js          Lógica del SIG + caché + sincronización
  estilos.css         Estilos de las marcas
  config.example.js   Plantilla de configuración (cópiala a config.js)
docs/
  DISTRIBUCION.md     Cómo repartir la extensión a los compañeros
```

> `extension/config.js` (con la URL real) **no se versiona** — está en
> `.gitignore`. En el repo solo va la plantilla `config.example.js`.

---

## Puesta en marcha

### 1) Apps Script (la API)

1. Ve a <https://sheets.google.com> y crea una hoja nueva. Ponle un nombre
   (p. ej. *Marcador KIMAK*). Decide en qué cuenta de Google vive — ver
   [Puntos a decidir](#puntos-a-decidir).
2. En esa hoja: menú **Extensiones → Apps Script**.
3. Borra el contenido de `Código.gs` y pega el de `apps-script/Codigo.gs`.
4. **Recomendado — pon una contraseña**: en la constante `TOKEN` (arriba del
   archivo) escribe una clave que compartirás solo con el equipo. Así la API
   solo responde a quien la sepa. Tendrás que poner exactamente la misma en
   `config.js` (`MK_TOKEN`). Ver [Contraseña vs. acceso](#contraseña-vs-acceso).
5. **Guarda** (💾).
6. **Implementar → Nueva implementación**.
   - Tipo: **Aplicación web**.
   - *Ejecutar como*: **Yo** (tu cuenta).
   - *Quién tiene acceso*: **Cualquier usuario** *(sin necesidad de iniciar
     sesión — así los compañeros no tienen login incómodo)*.
   - **Implementar**. La primera vez te pedirá autorizar permisos: acéptalos.
7. Copia la **URL de la aplicación web** (termina en `/exec`).

**Probar que responde**: pega esa URL en el navegador.
- Si **no** pusiste `TOKEN`: verás `{ "ok": true, "marcadas": [], "total": 0 }`.
- Si **sí** pusiste `TOKEN`: verás `{ "ok": false, "error": "token" }` — eso es
  buena señal (la API vive y la contraseña protege). Para probarla con clave,
  añade `?token=TU_CLAVE` al final de la URL.

Si ves cualquiera de las dos, el paso 1 está hecho.

> Cada vez que edites `Codigo.gs`, usa **Implementar → Gestionar
> implementaciones → (lápiz) → Versión: Nueva** para publicar los cambios sin
> que cambie la URL.

### 2) Extensión (conectarla a la API)

1. Copia `extension/config.example.js` a `extension/config.js`.
2. Abre `config.js` y pega tu URL en `MK_API_URL`.
   - Si activaste `TOKEN` en Apps Script, ponlo también en `MK_TOKEN`.
   - Opcional: pon tu nombre en `MK_USUARIO`.
3. Carga la extensión:
   - `chrome://extensions` → activa **Modo desarrollador**
     (arriba a la derecha).
   - **Cargar descomprimida** → selecciona la carpeta `extension/`.
4. Abre el SIG (`sig-design.kimak.com`) y prueba a marcar una fila. Deberías
   ver la nueva fila aparecer en la Google Sheet.

Si abres la hoja en dos ordenadores, lo que marque uno aparece en el otro en
cuanto pasa el intervalo de sincronización (30 s por defecto,
`MK_INTERVALO_MS`).

### 3) Distribución al equipo

Ver [`docs/DISTRIBUCION.md`](docs/DISTRIBUCION.md).

---

## Modo local (sin API)

Si `config.js` no tiene URL válida (sigue con `XXXX`), la extensión funciona
**solo en local** (como la versión original): guarda las marcas en
`chrome.storage.local` de ese PC. Útil para probar la parte visual sin montar
Google.

---

## Puntos a decidir

- **En qué cuenta de Google vive la hoja + script.** Recomendable una cuenta
  del departamento (no personal), para que no dependa de una persona.
- **Acceso de la web app.** Para que los compañeros no tengan que hacer login,
  al desplegar hay que elegir *Quién tiene acceso: **Cualquier usuario***. Es
  el ajuste que más confunde: si eliges "Solo yo" o "Usuarios de la
  organización", la extensión no podrá leer/escribir sin sesión.
- **Puente temporal vs. definitivo.** Se habló con el desarrollador del SIG de
  un marcador nativo. Si lo hace, esto es un puente; si va para largo, es la
  solución buena. Tenerlo en la balanza antes de invertir mucho.

## Contraseña vs. acceso

Son **dos capas distintas**, y la contraseña no sustituye al ajuste de Google:

| Capa | Qué controla | Dónde se pone |
|---|---|---|
| Acceso de Google ("Cualquier usuario") | Que la petición HTTPS *llegue* al script **sin pedir login de Google** | Al desplegar la web app |
| Contraseña (`TOKEN`) | Que el script solo *responda* a quien sabe la clave | En `Codigo.gs` + `config.js` |

Hay que dejar el acceso en **"Cualquier usuario"** sí o sí: la extensión llama
de forma anónima y, si Google exigiera login, no podría pasar (obligaría a un
flujo OAuth incómodo). Pero eso **no expone tus datos**: solo significa que la
puerta no pide carné de Google. Quien decide quién entra es la **contraseña**:

- Sin la clave → el script responde `{"ok":false,"error":"token"}`.
- Con la clave (va dentro de `config.js`, que solo tiene el equipo) → funciona.

Resultado: **quien tiene la contraseña tiene acceso; el resto, nada.** Es una
barrera ligera (la clave viaja en la petición), suficiente para un dato no
sensible como el número SIG.

## Seguridad y privacidad

- El único dato que sale es el **número SIG** (referencia sin datos sensibles);
  consultado con la empresa, se puede subir a Google Sheets.
- El repositorio es privado y **no** contiene la URL real, la contraseña ni
  rutas internas (`config.js` está en `.gitignore`).
