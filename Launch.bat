@echo off
title Word AI Add-in
cd /d "%~dp0"

echo.
echo ========================================
echo   Word AI Add-in - Starting...
echo ========================================
echo.

REM Install if needed
if not exist "node_modules" (
    echo Installing dependencies...
    call npm install
    if errorlevel 1 goto :error
    echo.
)

REM Build with localhost
echo Building...
set BASE_URL=http://localhost:3000
call npm run build
if errorlevel 1 goto :error
echo.

REM Start server and open folder
echo Starting server at http://localhost:3000
echo.
echo NEXT STEP: In Word - Insert -^> Add-ins -^> Upload My Add-in -^> select manifest.xml from the dist folder
echo.
echo KEEP THIS WINDOW OPEN while using the add-in.
echo Close it when done.
echo ========================================
echo.

start "" explorer "%~dp0dist"

npx --yes serve dist -p 3000 -s
goto :end

:error
echo.
echo Something went wrong. Make sure Node.js is installed: https://nodejs.org
pause

:end
