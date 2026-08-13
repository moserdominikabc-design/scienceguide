@echo off
setlocal
cd /d "%~dp0"

where py >nul 2>nul
if %errorlevel%==0 (
    set PY=py -3
) else (
    set PY=python
)

if not exist ".venv\Scripts\python.exe" (
    %PY% -m venv .venv || goto :error
)

.venv\Scripts\python.exe -m pip install --disable-pip-version-check -r requirements.txt || goto :error
.venv\Scripts\python.exe app.py
goto :eof

:error
echo.
echo Could not start ScienceGuide Excel Cleaner.
pause
exit /b 1
