@echo off
setlocal
cd /d "%~dp0"
echo === Comparador Excel CAF ===
echo.
python "%~dp0comparador.py"
echo.
echo --- Programa finalizado (codigo de salida: %errorlevel%) ---
echo Pulsa una tecla para cerrar...
pause >nul
endlocal
