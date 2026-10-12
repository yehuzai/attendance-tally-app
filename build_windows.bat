@echo off
rem Build AttendanceTally.exe for Windows. Run this ON a Windows machine.
setlocal
cd /d "%~dp0"

where py >nul 2>nul && (set PY=py -3) || (set PY=python)
%PY% -m venv .venv || goto :err
call .venv\Scripts\activate.bat
python -m pip install --upgrade pip
pip install -r requirements.txt pyinstaller || goto :err

pyinstaller --noconfirm --clean --windowed --onefile ^
  --name AttendanceTally ^
  app.py || goto :err

echo.
echo Done. Single file: dist\AttendanceTally.exe
goto :eof

:err
echo.
echo Build failed.
exit /b 1
