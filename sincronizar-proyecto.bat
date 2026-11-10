@echo off
setlocal
cd /d "%~dp0"

echo ========================================
echo  Sincronizar proyecto
echo ========================================
echo.

where node >nul 2>&1
if errorlevel 1 (
  echo ERROR: No se encontro Node.js.
  echo Instala Node.js LTS desde https://nodejs.org/ y vuelve a intentar.
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

if exist ".git\" (
  where git >nul 2>&1
  if errorlevel 1 (
    echo AVISO: Hay carpeta .git pero no se encontro Git en el PATH.
    echo Se omite git pull y se continua con npm install.
    echo.
  ) else (
    echo Actualizando codigo ^(git pull^)...
    git pull
    if errorlevel 1 (
      echo.
      echo ERROR: Fallo git pull. Revisa la conexion o el estado del repositorio.
      pause
      exit /b 1
    )
    echo.
  )
) else (
  echo No hay repositorio Git en esta carpeta.
  echo Se sincronizaran solo las dependencias npm.
  echo.
)

echo Sincronizando dependencias ^(npm install^)...
echo.
call npm install
if errorlevel 1 (
  echo.
  echo ERROR: Fallo npm install.
  pause
  exit /b 1
)

echo.
echo ========================================
echo  Sincronizacion completada.
echo  Ya puedes usar iniciar-app.bat
echo ========================================
echo.
pause
endlocal
