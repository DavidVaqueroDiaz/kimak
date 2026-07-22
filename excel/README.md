# AnidadoPRO — Calculadora de Tableros para Compras (Grupo KIMAK)

Macro VBA para Excel que calcula cuántos tableros hay que **pedir** por material
a partir del CSV de despiece exportado del SIG, usando nesting por franjas
(guillotina) que respeta la dirección de la veta de cada pieza.

Archivos:
- [`AnidadoPRO.bas`](AnidadoPRO.bas) — versión de tableros (V1).
- [`AnidadoPRO_V2.bas`](AnidadoPRO_V2.bas) — **V2**: lo mismo que la V1 **más un
  tercer cuadro con los metros de canto a pedir**. Ver [sección V2](#v2--cantos)
  al final. Es la versión recomendada; la V1 se conserva por compatibilidad.

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
  (la macro nunca une nada sola: decide siempre el usuario). En las filas
  amarillas, el desplegable muestra **primero los candidatos** detectados y
  después el resto de materiales (Excel no permite colorear los elementos de
  un desplegable, así que se ordenan para encontrarlos al instante).
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
- **M² REALES** (solo en la V2): metros cuadrados de tablero realmente
  consumidos = `EXACTO × área del tablero completo` (incluye piezas,
  separaciones y orilla). El último tablero parcial cuenta solo su fracción
  (0,1 → 0,1 × área), no un tablero entero. Con su total al final.
- Ambas tablas llevan **filtros** en la cabecera para ordenarlas por cualquier
  columna (material, tableros a pedir…).
- Una pieza que solo cabe usando el **tablero completo** (p. ej. 3050×1220 o
  3050×200: entra en el bruto pero no deja sitio a los 12 mm de orilla) sí se
  cuenta en los tableros a pedir y aparece avisada **en naranja** con su
  Cod. Pieza: va sin margen de orilla y hay que revisarla antes de cortar.
- Una pieza que no cabe ni en el tablero bruto aparece **en rojo** con su
  Cod. Pieza y cantidad, y no se cuenta: hay que resolverla antes del pedido
  (¿existe ese material en formato mayor? → cambia su medida y RECALCULAR).

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

## Materiales iguales escritos distinto

La macro agrupa como el **mismo material** los nombres que solo difieren en
guiones o espacios de más: `MDF--16`, `MDF---16` y `MDF   --16` se tratan como
uno solo automáticamente. Para diferencias mayores (una palabra escrita de otra
forma, p. ej. `BLANCA` vs `BLAN`) marca el posible duplicado en amarillo y se
unen a mano con **MISMO QUE**.

## Futuro previsto

La tabla MEDIDAS DE TABLERO es el punto de enganche para automatizar las
medidas por material desde un Excel/JSON maestro de materiales: cuando exista,
solo habrá que rellenarla automáticamente y saltarse la revisión manual, sin
tocar el algoritmo.

---

## <a name="v2--cantos"></a>V2 — Cantos

`AnidadoPRO_V2.bas` hace todo lo anterior y añade un tercer cuadro,
**CANTOS A PEDIR**, con los metros de canto a pedir por material. Se instala y
ejecuta igual que la V1 (macro `AnidadoPRO_V2`), con el mismo botón RECALCULAR
que recalcula los tres cuadros a la vez.

### Cómo calcula el canto

De la columna *Canteado* (`1L`, `2C 2L`, `1C 2L`, `0`, vacío):
- `nL` = nº de cantos en el lado **largo** (la medida **mayor** de la pieza).
- `nC` = nº de cantos en el lado **corto** (la medida **menor**).
- **Metros exactos por pieza** = `nL × ladoLargo + nC × ladoCorto`.
- **Metros a pedir** = exactos **+ 20 mm por cada canto** (borde), por pieza.
  Ej.: 10 piezas de 100×30 con `1L` → 1000 mm exactos / 1200 mm a pedir.

El canto no depende del nesting y se cuenta aunque la pieza no quepa en tablero.
La columna *Ingletado* **no** influye en el canto.

### Material del canto

- **Siempre** el material del tablero **sin el espesor**
  (`AGLOMERADO-S/ORDEN-16` → `AGLOMERADO-S/ORDEN`).
- Si la fila tiene *Comentarios Canteado*, ese texto se **añade entre
  paréntesis** y la celda se pinta de **naranja** para revisarla
  (`AGLOMERADO-S/ORDEN (CANTO PVC BLANCO)`).

Así cada canto queda siempre ligado a un material real. El cuadro tiene su
**propia columna MISMO QUE** (desplegable) y marca en amarillo los cantos casi
iguales: se unen a mano igual que los materiales de tablero y se pulsa
RECALCULAR. Las filas en naranja (con comentario) son las que conviene mirar
antes de pedir.

### Parámetros ajustables (al inicio del módulo)

| Constante | Valor | Qué es |
|---|---|---|
| `DEF_TAB_L` / `DEF_TAB_A` | 3050 / 1220 | Medida estándar de tablero (mm) |
| `MARGEN` | 12 | Orilla en los 4 lados (mm) |
| `SEP` | 17 | Separación entre piezas y franjas (mm) |
| `EXTRA_CANTO_MM` | 20 | Extra de canto por cada borde (mm) |
