@echo off
setlocal EnableExtensions DisableDelayedExpansion
cd /d "%~dp0" || exit /b 1

rem Use a reviewed local uv executable. See docs/installation.md.
set "UV_VERSION="
for /f "tokens=1,2" %%A in ('uv.exe --version 2^>nul') do if "%%A"=="uv" set "UV_VERSION=%%B"
if not "%UV_VERSION%"=="0.13.9" (
    echo Install uv 0.13.9 or select it on PATH. See docs/installation.md. 1>&2
    exit /b 1
)

set "UV_PROJECT_ENVIRONMENT=.venv"
uv.exe sync --locked --no-dev --no-editable --python 3.12.16
if errorlevel 1 exit /b 1
".venv\Scripts\python.exe" -I -c "from importlib.metadata import version; print('Installed outlook-classic-mcp', version('outlook-classic-mcp'))"
if errorlevel 1 exit /b 1
echo Server interpreter: "%CD%\.venv\Scripts\python.exe"
