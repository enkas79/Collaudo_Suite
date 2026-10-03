@echo off
setlocal

rem Build locale dell'eseguibile onedir. L'installer NSIS viene creato separatamente.
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

if errorlevel 1 exit /b %errorlevel%
echo Build completata in dist\CollaudoSuite
endlocal
