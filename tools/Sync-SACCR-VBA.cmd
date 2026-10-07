@echo off
setlocal EnableExtensions
set "REPO_ROOT=%~dp0.."
cd /d "%REPO_ROOT%"

set "SCRIPT=%~dp0Sync-SACCR-VBA.ps1"

if not exist "%SCRIPT%" (
    echo ERROR: PowerShell script not found:
    echo   %SCRIPT%
    echo.
    pause
    exit /b 1
)

set "WORKBOOK=%~1"

if not defined WORKBOOK (
    echo SACCR VBA synchronization
    echo.
    echo Enter the full path of the local SACCR .xlsm workbook.
    echo You can also drag an .xlsm file onto this .cmd file instead.
    echo.
    set /p "WORKBOOK=Workbook path: "
)

if not defined WORKBOOK (
    echo.
    echo No workbook path supplied. Nothing was changed.
    pause
    exit /b 1
)

if not exist "%WORKBOOK%" (
    echo.
    echo ERROR: Workbook not found:
    echo   %WORKBOOK%
    echo.
    pause
    exit /b 1
)

echo.
echo Repository:
echo   %CD%
echo Workbook:
echo   %WORKBOOK%
echo.
echo Synchronizing VBA source...
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" -WorkbookPath "%WORKBOOK%" -RepoRoot "%REPO_ROOT%"
set "RC=%ERRORLEVEL%"

echo.
if "%RC%"=="0" (
    echo SUCCESS: SACCR VBA synchronization completed.
) else (
    echo FAILED: SACCR VBA synchronization returned exit code %RC%.
)
echo.
pause
exit /b %RC%
