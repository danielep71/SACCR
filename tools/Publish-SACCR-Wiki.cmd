@echo off
setlocal
set "REPO_ROOT=%~dp0.."
cd /d "%REPO_ROOT%"
echo SA-CCR Benchmark Wiki publication
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Publish-SACCR-Wiki.ps1" -RepoRoot "%REPO_ROOT%"
set "EXIT_CODE=%ERRORLEVEL%"
echo.
if "%EXIT_CODE%"=="0" (echo Wiki publication completed successfully.) else (echo Wiki publication failed with exit code %EXIT_CODE%.)
pause
exit /b %EXIT_CODE%
