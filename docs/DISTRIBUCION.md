# Distribución a los compañeros

Objetivo: que cada persona del departamento tenga la extensión instalada y
apuntando a la **misma** Google Sheet, en unos 3 minutos.

## Preparar el paquete (lo hace una persona, una vez)

1. Asegúrate de tener `extension/config.js` creado y con la **URL real**
   (y `MK_TOKEN` si usáis token). Este archivo es el que lleva la configuración
   compartida.
2. Comprime la carpeta `extension/` en un ZIP. Debe incluir:
   ```
   manifest.json
   background.js
   content.js
   estilos.css
   config.js          ← con la URL real dentro
   ```
   > Ojo: `config.example.js` no hace falta en el ZIP; `config.js` sí.
3. Comparte ese ZIP por el canal interno habitual.

## Instalar (lo hace cada compañero)

1. Descomprime el ZIP en una **ruta fija en todos los equipos**, p. ej.
   `C:\marcador-kimak`.
   - Usar la **misma ruta** en todos los portátiles hace que Chrome asigne el
     **mismo ID de extensión** (todos usáis el mismo modelo de portátil / misma
     instalación). Útil si algún día queréis fijar el ID.
2. Abre `chrome://extensions`.
3. Activa **Modo desarrollador** (interruptor arriba a la derecha).
   - IT permite el modo desarrollador.
4. Pulsa **Cargar descomprimida** y selecciona la carpeta `C:\marcador-kimak`.
5. Abre el SIG (`sig-design.kimak.com`). Ya está: las marcas son compartidas.

## Actualizar la extensión más adelante

- Reparte un ZIP nuevo, que sobrescriba la carpeta `C:\marcador-kimak`.
- Cada compañero entra en `chrome://extensions` y pulsa el icono de
  **recargar** (🔄) sobre la extensión.
- Si solo cambió el código de Apps Script (no la extensión), **no** hay que
  repartir nada: basta publicar una versión nueva de la web app manteniendo la
  misma URL.

## ID fijo de la extensión (opcional)

Mientras se carga "descomprimida", el ID depende de la ruta de la carpeta. Con
la misma ruta en todos los equipos, el ID coincide. Si en el futuro necesitáis
un ID **garantizado e idéntico** independientemente de la ruta, se puede añadir
un campo `key` (clave pública) al `manifest.json`; se genera empaquetando la
extensión una vez (`chrome://extensions → Empaquetar extensión`) y copiando la
clave del `.pem` generado. No es necesario para el funcionamiento normal.

## Problemas típicos

| Síntoma | Causa probable | Solución |
|---|---|---|
| No aparecen los cuadros naranjas | Extensión no cargada o URL del SIG distinta | Revisa `chrome://extensions` y que la URL casa con `matches` del manifest |
| Marco pero no se comparte | `config.js` sin URL real, o web app mal desplegada | Abre la URL `/exec` en el navegador: debe dar `{"ok":true,...}` |
| `{"ok":false,"error":"token"}` | El `MK_TOKEN` de `config.js` no coincide con el `TOKEN` de Apps Script | Igualarlos |
| Marca y se revierte sola | La API no respondió (red / permisos de la web app) | Comprueba acceso "Cualquier usuario" en la implementación |
