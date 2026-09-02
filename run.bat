@echo off
setlocal
title Shift Priority App - Setup and Run

echo ============================================
echo   Checking for Node.js...
echo ============================================
where node >nul 2>nul
if %errorlevel% neq 0 (
    echo Node.js was not found on this system.
    echo.
    where winget >nul 2>nul
    if %errorlevel% neq 0 (
        echo winget is not available. Please install Node.js manually:
        echo   1. Go to https://nodejs.org
        echo   2. Download and install the LTS version
        echo   3. Re-run this file afterwards
        pause
        exit /b 1
    )
    echo Installing Node.js LTS via winget, please wait...
    winget install -e --id OpenJS.NodeJS.LTS --silent --accept-package-agreements --accept-source-agreements
    echo.
    echo Node.js was installed. Please CLOSE this window, open a NEW one,
    echo and double-click run.bat again so the system recognizes Node.
    pause
    exit /b 0
)

echo Node.js found:
node -v
echo.

echo ============================================
echo   Installing dependencies (first run only)...
echo ============================================
call npm install
if %errorlevel% neq 0 (
    echo.
    echo npm install failed. Check the error above.
    pause
    exit /b 1
)

echo.
echo ============================================
echo   Starting the app...
echo   A browser tab should open automatically.
echo   Press CTRL+C in this window to stop it.
echo ============================================
call npm run dev -- --open

pause
