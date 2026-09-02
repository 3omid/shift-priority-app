@echo off
setlocal
title Shift Priority - Push Update to GitHub

echo ============================================
echo   Sending your changes to GitHub...
echo ============================================
git add .
git commit -m "update"
if %errorlevel% neq 0 (
    echo.
    echo Nothing changed, or commit failed - check the message above.
)

git push
if %errorlevel% neq 0 (
    echo.
    echo Push failed. Copy the error above and send it back.
    pause
    exit /b 1
)

echo.
echo ============================================
echo   Done! Netlify will rebuild automatically
echo   in about 1-2 minutes.
echo ============================================
pause
