@echo off
setlocal

set ROOT_DIR=%~dp0
cd /d "%ROOT_DIR%"

where python >nul 2>nul
if errorlevel 1 (
  echo Python is not installed or not in PATH.
  echo Install Python 3.10+ from https://www.python.org/downloads/windows/
  pause
  exit /b 1
)

if not exist ".venv" (
  python -m venv .venv
)

call ".venv\Scripts\activate.bat"
python -m pip install --upgrade pip
pip install -r requirements.txt

if not exist "icon.ico" (
  python - <<'PY'
from PIL import Image
img = Image.open("icon.png")
img.save("icon.ico", sizes=[(16,16),(32,32),(48,48),(64,64),(128,128),(256,256)])
print("Created icon.ico")
PY
)

pyinstaller --noconfirm --clean --windowed --name "Lagardere" --icon icon.ico desktop_app.py

set ISCC=
if exist "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" set ISCC="C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
if exist "C:\Program Files\Inno Setup 6\ISCC.exe" set ISCC="C:\Program Files\Inno Setup 6\ISCC.exe"

if %ISCC%=="" (
  echo Inno Setup not found.
  echo Install Inno Setup 6 from https://jrsoftware.org/isdl.php
  pause
  exit /b 1
)

%ISCC% "Lagardere.iss"

echo Done. Installer is in: dist-installer\LagardereSetup.exe
pause
