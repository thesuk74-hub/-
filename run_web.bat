@echo off
setlocal
cd /d "%~dp0"
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8
title Construction Period Review - web server

rem ---- find Python (prefer the "py" launcher; avoid the Microsoft Store stub) ----
set "PY="
py -3 -c "import sys" >nul 2>nul && set "PY=py -3"
if not defined PY python -c "import sys" >nul 2>nul && set "PY=python"
if not defined PY goto nopython
%PY% -c "import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)" || goto oldpython
echo [1/3] Python found:
%PY% --version

rem ---- install packages ----
echo [2/3] Checking required packages ...
%PY% -m pip install -q --disable-pip-version-check -r requirements.txt
if errorlevel 1 goto piperror

rem ---- start server; the browser opens automatically when the server is ready ----
echo [3/3] Starting server: http://127.0.0.1:8000
echo       Keep this window open while using the program. To stop: press Ctrl+C or close this window.
echo.
%PY% -m cpa web --host 127.0.0.1 --port 8000 --open
echo.
echo [Server stopped] If an error message is shown above, please send a screenshot of this window.
pause
exit /b 0

:nopython
echo [ERROR] Python was not found.
echo   1. Install Python 3.10 or later from https://www.python.org/downloads/
echo   2. In the installer, check "Add python.exe to PATH".
echo   3. Run this file again.
pause
exit /b 1

:oldpython
echo [ERROR] Python 3.10 or later is required.
%PY% --version
pause
exit /b 1

:piperror
echo [ERROR] Package installation failed. Check the internet connection or company proxy settings.
echo   Manual install:  %PY% -m pip install -r requirements.txt
pause
exit /b 1
