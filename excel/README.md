# AnidadoPRO — Calculadora de Tableros para Compras (Grupo KIMAK)

Macro VBA para Excel que calcula cuántos tableros hay que **pedir** por material
a partir del CSV de despiece exportado del SIG, usando nesting por franjas
(guillotina) que respeta la dirección de la veta de cada pieza.

Archivo: [`AnidadoPRO.bas`](AnidadoPRO.bas)

---

## Instalación (una sola vez por usuario)

1. Abrir Excel y pulsar `Alt + F11` (editor de VBA).
2. En el panel izquierdo, buscar **VBAProject (PERSONAL.XLSB)**.
   - Si no existe: grabar una macro vacía cualquiera (Vista → Macros → Grabar
     macro → guardar en *Libro de macros personal*) y detenerla. Eso crea el
     libro personal.
3. Clic derecho sobre PERSONAL.XLSB → **Importar archivo…** → seleccionar
   `AnidadoPRO.bas`.
4. Cerrar el editor. La macro `AnidadoPRO` queda disponible en cualquier
   libro que se abra (`Alt + F8`).

> Consejo: añadirla a la barra de acceso rápido para lanzarla con un clic
> (Archivo → Opciones → Barra de herramientas de acceso rápido → Macros).

## Uso

### Modo 1 — un solo CSV
1. Abrir el CSV del SIG en Excel.
2. `Alt + F8` → `AnidadoPRO` → Ejecutar.

### Modo 2 — pedido con varios CSV
1. Ejecutar `AnidadoPRO` **sin datos delante** (libro vacío o Excel
   recién abierto).
2. La macro pregunta si quieres importar: se abre el explorador de Windows.
3. Seleccionar todos los CSV del pedido (`Ctrl + clic`) y aceptar.
4. La macro los combina en un libro nuevo y calcula todo el pedido junto.

## Qué genera

Debajo de los datos aparecen dos tablas:

**MEDIDAS DE TABLERO** — una fila por material, precargada con 3050 × 1220 mm.
- Editar largo/ancho de los materiales cuyo tablero mida distinto.
- Columna **MISMO QUE**: si dos materiales son en realidad el mismo (error de
  escritura en el diseño), elegir en el desplegable el material "bueno"; al
  recalcular, sus piezas se anidan juntas. Los nombres casi idénticos con el
  mismo espesor se marcan en **amarillo** como aviso de posible duplicado
  (la macro nunca une nada sola: decide siempre el usuario).
- Botón **RECALCULAR**: regenera el resultado al instante con las medidas y
  uniones introducidas, sin volver a preguntar nada.

**TABLEROS A PEDIR** — el resultado para compras:

| MATERIAL | MEDIDA TABLERO | EXACTO | A PEDIR |
|---|---|---|---|
| FR MDF-MEL_BLANCA SR209 SOFTIII-16 | 3050 x 1220 | 1.3 | **2** |
| … | … | … | … |
| TOTAL TABLEROS | | | **8** |

- **A PEDIR** es la cantidad a comprar (redondeo hacia arriba, en verde).
- **EXACTO** es el cálculo con un decimal, como referencia de aprovechamiento.
- Si una pieza no cabe en ningún tablero, aparece una fila **en rojo** con sus
  dimensiones: hay que revisarla antes de hacer el pedido.

## Reglas del cálculo

| Parámetro | Valor |
|---|---|
| Tablero por defecto | 3050 × 1220 mm (editable por material) |
| Orilla | 12 mm en los 4 lados |
| Separación entre piezas y franjas | 17 mm (12 libre + 5 kerf) |

Reglas de veta (columna *Dir. Veta* del CSV):
- `Largo`, `Corto` o cualquier otro texto → la pieza no gira; la medida de la
  columna Largo (que es donde va la veta) se alinea con el largo del tablero.
  El SIG escribe `Corto` cuando la veta va en el largo pero el largo es la
  medida pequeña de la pieza.
- `Ancho` → la pieza no gira; se coloca girada 90° (caso raro, contemplado).
- Vacío o `0` → sin veta; gira solo si *Permite giro* = 1.
- Si una fila trae veta y *Permite giro* = 1 a la vez, **la veta manda**: no gira.

## Futuro previsto

La tabla MEDIDAS DE TABLERO es el punto de enganche para automatizar las
medidas por material desde un Excel/JSON maestro de materiales: cuando exista,
solo habrá que rellenarla automáticamente y saltarse la revisión manual, sin
tocar el algoritmo.
