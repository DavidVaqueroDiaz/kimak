# Piezas CNC — copiar una sola copia de cada pieza

Dos scripts de Windows (PowerShell + ventana) para, a partir de un índice JSON
de una carpeta, dejar en una carpeta destino **un único archivo por pieza**: el
más reciente.

| Archivo | Para qué |
|---|---|
| `Generar-Indice.bat` / `.ps1` | Recorre una carpeta (aunque sea `\\servidor10\...`) y crea el JSON. Solo lo necesitas si no tienes ya el índice. |
| `Copiar-Piezas.bat` / `.ps1` | **El principal.** Lee el JSON, agrupa las piezas repetidas y copia la más reciente de cada una a la carpeta que escribas. |

Se abren con doble clic en el `.bat` (lleva `-ExecutionPolicy Bypass`, así no hay
que tocar la política de PowerShell del PC).

## Qué hace exactamente `Copiar-Piezas`

1. Lee el JSON y busca en el **nombre** de cada archivo el código de pieza:
   `1002275-001-001_A`, `902275-001-001_A`, `943625-001-001-B`, `000875483-001-001_A`…
   Lo busca en cualquier parte del nombre, así que también valen
   `PIEZA 10_943251-001-001_A MDF 16mm 1und.ard` o `[NUEVO] 878253-003-004_A_HPL natural`.
2. Todo lo que venga **detrás del código es la misma pieza** (material, uds., programa distinto).
3. Agrupa y de cada grupo copia **solo el de fecha más cercana a hoy**.
4. Escribe un informe `_informe_piezas_AAAAMMDD_HHMMSS.csv` en la carpeta destino
   con **todas** las filas: lo copiado, lo descartado y por qué.

### Reglas de agrupación (lo que cuenta como "la misma pieza")

| Caso | Decisión |
|---|---|
| Revisión `_A`, `_B`, `_D` | **Piezas distintas.** El código completo incluye la letra. En el índice de M2 hay 416 piezas con varias revisiones. |
| `Cara A` / `Cara B` | **Se copian por separado** (casilla activada). Son dos programas distintos de la misma pieza: agrupándolos se perdería el de la 2ª cara. Hay 130 códigos afectados. |
| Extensión distinta (`.cix`, `.hop`) | No se copian: por defecto solo se procesan `.ard`. Si las añades en el cuadro de extensiones, cuentan como piezas distintas (son máquinas distintas). |
| Archivos sin código (`DM 25mm.`, `PLACAS PVC ESPUMADO 10 MM BLANCO`) | **No se copian.** Son nestings/tableros. Salen en el CSV como `SIN CODIGO`. |

`cara buena` y `cara mala` **no** se confunden con cara A/B (son notas de
orientación, 890 archivos en el índice de M2).

## Opciones de la ventana

- **Extensiones**: por defecto **solo `.ard`**. Si algún día quieres los de otras máquinas, escribe `.ard;.cix;.hop`.
- **Filtro de ruta** (opcional): solo procesa archivos cuya ruta contenga ese texto (p. ej. `ZARA HOME`).
- **Separar Cara A / Cara B**: activado.
- **Comprobar fechas reales en disco**: solo mira el disco en los grupos con duplicados, para desempatar con la fecha real en lugar de la del índice. Más lento pero más fiable si el índice es viejo.
- **Tener en cuenta también la fecha de creación**: desactivado. Ojo: al copiar un archivo su fecha de creación pasa a ser la del día de la copia, así que una copia reciente de un programa viejo ganaría a un programa realmente más nuevo. Por eso por defecto solo se usa la fecha de modificación.
- **Sobrescribir**: si no se marca, lo que ya esté en destino se respeta (`OMITIDO`), así se puede relanzar sin duplicar.
- **SIMULAR**: no copia nada, solo genera el CSV. **Úsalo la primera vez.**

## Detalles que te van a importar

- **Nombres**: se conserva el nombre original. Si dos archivos distintos fueran a llamarse igual, el segundo se guarda como `nombre_2.ard` (no se pisa nada).
- **Si el más reciente ya no existe** en su ruta, coge el siguiente más reciente que sí exista, y lo deja anotado en el CSV.
- **Rutas largas**: soporta rutas de más de 260 caracteres (en el índice de M2 hay 1.478 archivos así; con `Copy-Item` normal fallarían).
- **Empates de fecha**: gana la ruta más corta (suele ser la carpeta original, no una copia).
- El índice de M2 completo (41.053 archivos) tarda unos 20–60 s en clasificarse, más el tiempo de copia por red.

## Uso sin ventana (tarea programada)

```powershell
.\Copiar-Piezas.ps1 -SinInterfaz -Indice "C:\indices\M2.json" -Destino "D:\PIEZAS" -Simular
.\Generar-Indice.ps1 -SinInterfaz -Carpeta "\\servidor10\Intercambio_General\CNC M2\ZARA" -Salida "C:\indices\zara.json"
```

Parámetros de `Copiar-Piezas.ps1`: `-Extensiones`, `-FiltroRuta`, `-NoSepararCaras`,
`-FechasDeDisco`, `-UsarFechaCreacion`, `-Sobrescribir`, `-Simular`.

## Vía rápida sin tocar PowerShell (`COPIAR_PIEZAS.bat`)

Si ya tienes la lista filtrada (`piezas_a_copiar.json`, un objeto por pieza con
`ruta_completa` y `nombre_destino`):

1. Deja `COPIAR_PIEZAS.bat` y `piezas_a_copiar.json` **en la misma carpeta**.
2. Abre el `.bat` con el Bloc de notas y cambia la línea `set "DESTINO=D:\PIEZAS"`.
3. Guarda y doble clic.

Son 20 líneas de texto ASCII: no hay que cambiar la política de PowerShell, soporta
rutas de más de 260 caracteres y va diciendo por dónde va.

## Usar ChatGPT para generar la lista (`PROMPT-ChatGPT.md`)

**ChatGPT no puede copiar los archivos**: su entorno no ve `\\servidor10\...`. Lo que sí
puede es leer el índice de una carpeta, aplicar las reglas y devolverte
`piezas_a_copiar.json` + `COPIAR_PIEZAS.bat` para que los ejecutes tú. El prompt de
`PROMPT-ChatGPT.md` lleva las reglas escritas una a una para que el resultado sea el
mismo que el de estos scripts.

## Formato del índice JSON

Array plano; el script solo necesita `nombre`, `ruta_completa` y `fecha_modificacion`
(`extension`, `nombre_sin_ext`, `ruta` y `fecha_creacion` son opcionales):

```json
[{"nombre":"971205-001-001_A DM 25_1ud.ard","nombre_sin_ext":"971205-001-001_A DM 25_1ud",
  "extension":".ard","servidor":"M2","ruta":"\\\\servidor10\\...\\971196-1",
  "ruta_completa":"\\\\servidor10\\...\\971196-1\\971205-001-001_A DM 25_1ud.ard",
  "fecha_modificacion":"2026-07-10T10:38:43.9500889+02:00"}]
```
