@echo off
chcp 65001 > nul
title Procesador de Facturas

echo.
echo  =========================================
echo   PROCESADOR DE FACTURAS PUBLICITARIAS
echo  =========================================
echo.

cd /d "%~dp0"

"C:\Users\franco.faustin\AppData\Local\Python\pythoncore-3.14-64\python.exe" procesar_facturas.py

echo.
echo  =========================================
echo   Proceso finalizado. Presiona cualquier
echo   tecla para cerrar esta ventana.
echo  =========================================
pause > nul
