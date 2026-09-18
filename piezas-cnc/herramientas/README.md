# Herramientas de agrupación por material

Las ejecuta Claude, no hacen falta en el PC del taller. Sirven para regenerar el
documento de piezas agrupadas por material cuando haya un índice nuevo.

| Archivo | Qué hace |
|---|---|
| `clasifica.py` | Lee el texto que va **detrás del código** en el nombre del archivo y deduce material y espesor. Quita unidades (`_1und`, `2uds`), notas entre paréntesis/corchetes, `cara A/B` y `cara buena/mala`. 40 familias de material por palabras clave. |
| `gen_docx.js` | Con `materiales.json` genera `piezas_por_material.docx`: portada, índice de materiales y un bloque por material con título en negrita a 20 pt y salto de página. Necesita `npm install docx`. |

Cadena completa: `piezas_a_copiar.json` → `clasifica.py` → `materiales.json` → `gen_docx.js` → `.docx`

## Lo que NO puede adivinar

- 1.362 piezas se llaman solo con el código, sin material: van a un bloque aparte.
- 93 tienen un texto que no es un material reconocible (`mod`, `techo`, `frente`).
- DM y MDF se agrupan juntos (es el mismo tablero escrito de dos formas).
- El espesor es el primer número plausible (0,5–120) del texto: `+2mm pvc19` lo lee como 2.
