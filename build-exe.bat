@echo off
setlocal
title Shift Priority App - Build Windows EXE

where node >nul 2>nul
if %errorlevel% neq 0 (
    echo Node.js was not found. Run run.bat first to install it, then come back here.
    pause
    exit /b 1
)

echo Node version:
node -v
echo.

echo ============================================
echo   Installing dependencies (Electron included)
echo   This can take several minutes the first time.
echo ============================================
call npm install
if %errorlevel% neq 0 (
    echo.
    echo npm install FAILED. Copy the error above and send it back.
    pause
    exit /b 1
)

echo.
echo ============================================
echo   Building the app bundle...
echo ============================================
call npm run build
if %errorlevel% neq 0 (
    echo.
    echo "npm run build" FAILED. Copy the error above and send it back.
    pause
    exit /b 1
)

echo.
echo ============================================
echo   Packaging the Windows .exe installer...
echo ============================================
set CSC_IDENTITY_AUTO_DISCOVERY=false
call npx electron-builder --win
if %errorlevel% neq 0 (
    echo.
    echo electron-builder FAILED. Copy the FULL error above and send it back.
    echo Common causes: no internet access, or antivirus blocking the download.
    pause
    exit /b 1
)

echo.
if exist release (
    echo SUCCESS. Your installer is inside the "release" folder:
    dir /b release\*.exe
    explorer release
) else (
    echo The build finished without an error, but no "release" folder was found.
    echo Please copy everything printed above and send it back.
)

pause
