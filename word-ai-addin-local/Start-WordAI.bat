@echo off
title Word AI Add-in - Local Server
echo.
echo ========================================
echo   Word AI Add-in - Local Server
echo ========================================
echo.
echo Starting local server...
echo.
echo KEEP THIS WINDOW OPEN while using the add-in.
echo.
echo When the server is ready:
echo   1. Open Microsoft Word
echo   2. Go to Insert -> Add-ins -> Upload My Add-in
echo   3. Select the manifest.xml file in this folder
echo   4. Use AI Assistant from the Home tab
echo.
echo To stop: Close this window.
echo ========================================
echo.

REM Try Node.js first (npx serve)
where npx >nul 2>&1
if %errorlevel% equ 0 (
  echo Using Node.js...
  npx --yes serve -p 3000 -s
  goto :end
)

REM Fall back to Python
where python >nul 2>&1
if %errorlevel% equ 0 (
  echo Using Python...
  python -m http.server 3000 -d .
  goto :end
)

echo ERROR: Neither Node.js nor Python was found.
echo.
echo Please install one of these:
echo   - Node.js: https://nodejs.org
echo   - Python:  https://python.org
echo.
pause
:end
