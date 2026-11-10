@echo off
setlocal
cd /d "%~dp0"

echo ========================================
echo  Nomina - Relacion / Pago Terceros
echo ========================================
echo.

where node >nul 2>&1
if errorlevel 1 (
  echo ERROR: No se encontro Node.js.
  echo Instala Node.js LTS desde https://nodejs.org/ y vuelve a ejecutar este archivo.
  echo.
  pause
  exit /b 1
)

where npm >nul 2>&1
if errorlevel 1 (
  echo ERROR: No se encontro npm.
  echo Reinstala Node.js LTS desde https://nodejs.org/ e intenta de nuevo.
  echo.
  pause
  exit /b 1
)

echo Node:
node -v
echo npm:
npm -v
echo.

if not exist "node_modules\" (
  echo Instalando dependencias ^(npm install^)... Esto puede tardar unos minutos.
  echo.
  call npm install
  if errorlevel 1 (
    echo.
    echo ERROR: Fallo npm install.
    pause
    exit /b 1
  )
  echo.
)

echo Abriendo http://localhost:5173 en el navegador...
start "" "http://localhost:5173"

echo.
echo Iniciando la app. Deja esta ventana abierta mientras la uses.
echo Cierra la ventana para detener el servidor.
echo.

call npm run dev

echo.
pause
endlocal
