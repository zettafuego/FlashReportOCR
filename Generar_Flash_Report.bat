@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ============================================
echo   FLASH REPORT - Minera Casma
echo   Foto de manifiesto -^> Excel oficial
echo ============================================
echo.

if "%~1"=="" (
  echo Arrastra una o mas fotos sobre este .bat
  echo o ejecuta:
  echo   Generar_Flash_Report.bat "C:\ruta\foto.jpg"
  echo.
  echo Tambien puedes usar:
  echo   python flash_report.py --json ejemplos\content_jeferson.json --no-gui
  echo.
  pause
  exit /b 1
)

python flash_report.py %*
if errorlevel 1 (
  echo.
  echo Hubo un error. Revisa el mensaje anterior.
  pause
  exit /b 1
)

echo.
echo Listo. Revisa la carpeta "salidas".
pause
