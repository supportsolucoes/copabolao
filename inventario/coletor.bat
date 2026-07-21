@echo off
title Inventario de Maquinas
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0coletor.ps1"
echo.
pause
