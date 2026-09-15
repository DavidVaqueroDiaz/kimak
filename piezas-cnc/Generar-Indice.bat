@echo off
setlocal
title KIMAK - Generar indice
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0Generar-Indice.ps1"
if errorlevel 1 pause
