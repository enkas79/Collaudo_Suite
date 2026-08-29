@echo off
setlocal
cd /d "%~dp0"

if not exist .venv (
    py -3 -m venv .venv
)
call .venv\Scripts\activate.bat
python -m pip install --upgrade pip
python -m pip install -r requirements.txt pyinstaller

python -m PyInstaller --noconfirm --clean --windowed ^
  --name "CollaudoSuite" ^
  --collect-all matplotlib ^
  --collect-all nltk ^
  --hidden-import PySide6.QtPrintSupport ^
  --hidden-import collaudo_suite.checklist.data ^
  --add-data "collaudo_suite\checklist\data;collaudo_suite\checklist\data" ^
  --add-data "version.txt;." ^
  run_suite.py

if errorlevel 1 (
    echo.
    echo Compilazione non riuscita.
    pause
    exit /b 1
)

echo.
echo Eseguibile creato in: dist\CollaudoSuite\CollaudoSuite.exe
pause
