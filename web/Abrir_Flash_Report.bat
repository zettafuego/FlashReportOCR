@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ============================================
echo  FLASH REPORT - Minera Casma
echo  http://127.0.0.1:8765/index.html
echo ============================================
echo.

where python >nul 2>&1
if errorlevel 1 (
  echo Python no encontrado. Abriendo archivo local...
  start "" "%~dp0index.html"
  pause
  exit /b 0
)

REM Liberar puerto 8765 si estaba ocupado
for /f "tokens=5" %%a in ('netstat -ano 2^>nul ^| findstr ":8765" ^| findstr "LISTENING"') do (
  taskkill /F /PID %%a >nul 2>&1
)

start "FlashReport-Server" /min python -m http.server 8765
timeout /t 2 /nobreak >nul
start "" http://127.0.0.1:8765/index.html

echo.
echo App abierta. Deja la ventana minimizada del servidor abierta.
echo Para cerrar el servidor: cierra la ventana "FlashReport-Server".
echo.
pause
