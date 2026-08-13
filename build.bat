@echo off
REM ============================================================
REM  Unimex Customs Summary - Build Script
REM  Run this once on your Windows machine to produce the .exe.
REM  Requires Python 3.10+ already installed.
REM ============================================================

REM --- Pick a Python interpreter -----------------------------
REM  Prefer the Windows launcher "py -3": it resolves the real
REM  installed Python from the registry and ignores whatever
REM  happens to be first on PATH (e.g. an unrelated venv). Fall
REM  back to plain "python" only if the launcher isn't present.
where py >nul 2>nul && (set "PY=py -3") || (set "PY=python")

echo.
echo === Using interpreter ===
%PY% --version
if errorlevel 1 (
    echo.
    echo [ERROR] No usable Python found. Install Python 3.10+ from
    echo         python.org and make sure "py" or "python" works in a
    echo         new terminal, then run this again.
    pause
    exit /b 1
)

echo.
echo === Installing required packages ===
%PY% -m pip install --upgrade pip
%PY% -m pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo [ERROR] pip install failed. Make sure Python is installed and on your PATH.
    pause
    exit /b 1
)

echo.
echo === Building executable with PyInstaller ===
REM --onefile    : produce a single .exe instead of a folder
REM --console    : keep the console window (we want users to see output)
REM --name       : sets the output filename (UnimexCustoms.exe)
REM --clean      : wipe PyInstaller caches first to avoid stale builds
REM --noupx      : do not UPX-pack (UPX-packed exes trip more antivirus engines,
REM                which matters now that the exe self-downloads updates)
%PY% -m PyInstaller --onefile --console --clean --noupx --name UnimexCustoms customs_processor.py
if errorlevel 1 (
    echo.
    echo [ERROR] PyInstaller failed. See messages above.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  Build complete!
echo.
echo  The .exe is at:
echo    dist\UnimexCustoms.exe
echo.
echo  To distribute: copy the .exe to a user-writable folder on
echo  the customs team's machine, run it once to create the
echo  folders, then put the parts database in database\.
echo ============================================================
echo.
pause
