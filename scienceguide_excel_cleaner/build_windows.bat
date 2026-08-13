@echo off
setlocal
cd /d "%~dp0"

where py >nul 2>nul
if %errorlevel%==0 (
    set PY=py -3
) else (
    set PY=python
)

if not exist ".venv-build\Scripts\python.exe" (
    %PY% -m venv .venv-build || goto :error
)

.venv-build\Scripts\python.exe -m pip install --upgrade pip || goto :error
.venv-build\Scripts\python.exe -m pip install -r requirements-build.txt || goto :error
.venv-build\Scripts\pyinstaller.exe --noconfirm --clean --windowed --name "ScienceGuide Excel Cleaner" app.py || goto :error

echo.
echo Build complete. See dist\ScienceGuide Excel Cleaner\
pause
goto :eof

:error
echo.
echo Build failed.
pause
exit /b 1
