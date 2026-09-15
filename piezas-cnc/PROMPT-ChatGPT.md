Actúa como analista de datos. Te adjunto un índice en JSON con los archivos de una
carpeta de red de un taller CNC: cada archivo es un programa de Alphacam de una pieza.
Necesito quedarme con UNA sola copia de cada pieza.

Procesa el adjunto ENTERO con código Python (pandas/json). No lo leas "a ojo" ni lo
resumas por muestreo.

=== FORMATO DEL ADJUNTO ===
Array JSON de objetos con estos campos:
  nombre, nombre_sin_ext, extension, servidor, ruta, ruta_completa, fecha_modificacion
fecha_modificacion va en ISO 8601 con zona horaria (ej. 2026-07-10T10:38:43.9500889+02:00).

=== REGLAS (aplícalas literalmente) ===
1. Quédate SOLO con los archivos de extensión .ard. Descarta cualquier otra extensión.

2. Busca el código de pieza dentro de "nombre_sin_ext" con este regex de Python
   (puede estar en cualquier parte del nombre, no solo al principio):
       (?<![0-9])(\d{6,9})-(\d{3,4})-(\d{2,4})(?:[_-]([A-Za-z]))?(?![0-9A-Za-z])

3. Los archivos SIN código no se copian: son nestings o tableros ("DM 25mm.",
   "PLACAS PVC ESPUMADO 10 MM BLANCO"). Cuéntalos aparte y dime cuántos son.

4. Normaliza el código: a cada uno de los tres grupos de dígitos quítale los ceros de
   la izquierda y rellénalo con ceros hasta 6, 3 y 3 dígitos. La letra final es la
   revisión, va en mayúscula y FORMA PARTE del código:
   881048-001-001_A y 881048-001-001_D son PIEZAS DISTINTAS, se copian las dos.

5. Detecta la cara del programa en "nombre_sin_ext" (case-insensitive):
       cara B ->  cara\s*[_\-\.]?\s*b(?![a-zA-Z])
       cara A ->  cara\s*[_\-\.]?\s*a(?![a-zA-Z])
   OJO: "cara buena" y "cara mala" NO son cara A ni cara B, son notas de orientación.

6. Agrupa por (codigo, cara, extension) y quédate con el de fecha_modificacion MÁS
   RECIENTE de cada grupo. Si hay empate de fecha, el de ruta_completa más corta.

7. El nombre en destino es el nombre original del archivo. Si dos piezas distintas
   fueran a llamarse igual, al segundo le añades _2, al tercero _3, etc.

=== QUÉ QUIERO QUE ME DEVUELVAS ===
A) Un resumen en pantalla: archivos totales, con código, sin código, piezas distintas,
   archivos descartados por duplicados, y las 10 piezas que más veces se repiten.

B) Un archivo "piezas_a_copiar.json" para descargar: array JSON, codificado en UTF-8
   SIN BOM, con un objeto por pieza y estos campos:
       codigo_pieza, cara, nombre, nombre_destino, extension, ruta_completa,
       fecha_modificacion, copias_descartadas

C) Un archivo "COPIAR_PIEZAS.bat" para descargar, SOLO con caracteres ASCII y saltos de
   línea CRLF, con EXACTAMENTE este contenido, cambiando únicamente la ruta de la línea
   DESTINO por la carpeta que te indico al final:

--- INICIO DEL .BAT ---
@echo off
setlocal
rem ==============================================================
rem   KIMAK - Copiar las piezas ya filtradas (una copia por pieza)
rem
rem   1) Deja este archivo y piezas_a_copiar.json en la misma carpeta
rem   2) Cambia la ruta de la linea DESTINO por la que quieras
rem   3) Guarda y haz doble clic en este archivo
rem ==============================================================

set "DESTINO=D:\PIEZAS"

set "AQUI=%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -Command "& { $d=[IO.Path]::GetFullPath($env:DESTINO); [IO.Directory]::CreateDirectory($d)|Out-Null; $f=Join-Path $env:AQUI 'piezas_a_copiar.json'; if(-not (Test-Path -LiteralPath $f)){ Write-Host ('NO ENCUENTRO '+$f); exit 1 }; $L=Get-Content -LiteralPath $f -Raw -Encoding UTF8 | ConvertFrom-Json; $n=0;$e=0;$i=0;$t=$L.Count; Write-Host ('Copiando {0} piezas a {1}' -f $t,$d); Write-Host ''; foreach($p in $L){ $i++; $o=$p.ruta_completa; if($o.StartsWith('\\')){$o='\\?\UNC\'+$o.Substring(2)}elseif($o -match '^[A-Za-z]:\\'){$o='\\?\'+$o}; $x=Join-Path $d $p.nombre_destino; if($x.StartsWith('\\')){$x='\\?\UNC\'+$x.Substring(2)}elseif($x -match '^[A-Za-z]:\\'){$x='\\?\'+$x}; try{ [IO.File]::Copy($o,$x,$true); $n++ }catch{ $e++; if($e -le 15){ Write-Host ('  NO SE PUDO COPIAR: '+$p.nombre) } }; if(($i %% 250) -eq 0){ Write-Host ('  {0} / {1}' -f $i,$t) } }; Write-Host ''; Write-Host ('TERMINADO.  Copiadas: {0}   Errores: {1}' -f $n,$e) }"
echo.
pause
--- FIN DEL .BAT ---

=== IMPORTANTE ===
Tú NO tienes acceso a la red del taller (\\servidor10\...), así que no intentes copiar,
abrir ni verificar los archivos originales: solo generas los dos archivos para descargar.
Los ejecuto yo en mi PC.

=== CARPETA DONDE QUIERO LAS COPIAS ===
>>> ESCRIBE AQUÍ LA RUTA, por ejemplo: D:\PIEZAS\32610690007 <<<
