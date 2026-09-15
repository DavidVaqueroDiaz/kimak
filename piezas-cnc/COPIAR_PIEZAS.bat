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
