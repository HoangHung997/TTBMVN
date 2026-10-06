@echo off
setlocal
title Cai dat TTBMVN Excel Tools

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install-TTBMVN.ps1" -TrustPilotCertificate
set "exitCode=%ERRORLEVEL%"

echo.
if not "%exitCode%"=="0" (
    echo Cai dat khong thanh cong. Hay doc thong bao loi phia tren.
) else (
    echo Cai dat thanh cong. Ban co the mo lai Excel.
)
echo.
pause
exit /b %exitCode%
