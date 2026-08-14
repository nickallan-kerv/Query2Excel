@echo off
setlocal

REM Resolve repository root from this script location.
set "REPO_ROOT=%~dp0"
set "WEB_PROJECT=%REPO_ROOT%Query2Excel.Web\Query2Excel.Web.csproj"

if not exist "%WEB_PROJECT%" (
    echo Could not find web project:
    echo %WEB_PROJECT%
    pause
    exit /b 1
)

where dotnet >nul 2>nul
if errorlevel 1 (
    echo dotnet was not found on PATH. Install .NET SDK to run this launcher.
    pause
    exit /b 1
)

echo Starting Query2Excel.Web...
echo Project: %WEB_PROJECT%
echo.

dotnet run --project "%WEB_PROJECT%" -- %*

if errorlevel 1 (
    echo.
    echo Web app exited with an error.
    pause
    exit /b 1
)

endlocal