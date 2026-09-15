@echo off
setlocal
title KIMAK - Copiar piezas
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0Copiar-Piezas.ps1"
if errorlevel 1 pause
